package ru.objectsfill.style;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.dhatim.fastexcel.reader.CellAddress;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import ru.objectsfill.exception.ExcelExportException;

import javax.xml.parsers.SAXParserFactory;
import java.io.InputStream;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * Читает стили ячеек из Excel-шаблона потоково (SAX), без загрузки всего файла в память.
 * Обрабатывает не более {@code rowLimit} строк на каждом листе.
 */
public class PoiStyleReader {

    public static final int DEFAULT_STYLE_READ_DEPTH = 100;

    private final int rowLimit;

    public PoiStyleReader() {
        this(DEFAULT_STYLE_READ_DEPTH);
    }

    public PoiStyleReader(int rowLimit) {
        this.rowLimit = rowLimit;
    }

    /**
     * @param templateStream поток Excel-шаблона
     * @return карта: имя листа → (адрес ячейки → атрибуты стиля)
     */
    public Map<String, Map<CellAddress, CellStyleAttributes>> read(InputStream templateStream) {
        try (OPCPackage pkg = OPCPackage.open(templateStream)) {
            XSSFReader xssfReader = new XSSFReader(pkg);
            StylesTable stylesTable = xssfReader.getStylesTable();

            Map<String, Map<CellAddress, CellStyleAttributes>> result = new LinkedHashMap<>();
            XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();

            while (iter.hasNext()) {
                try (InputStream sheetStream = iter.next()) {
                    String sheetName = iter.getSheetName();
                    result.put(sheetName, parseSheet(sheetStream, stylesTable));
                }
            }
            return result;
        } catch (ExcelExportException e) {
            throw e;
        } catch (Exception e) {
            throw new ExcelExportException("Failed to read styles from template", e);
        }
    }

    /**
     * Потоково разбирает XML одного листа и собирает стили ячеек.
     * Останавливается по достижении {@code rowLimit} строк.
     *
     * @param sheetStream поток XML-данных листа
     * @param stylesTable таблица стилей из POI
     * @return карта: адрес ячейки → атрибуты стиля
     */
    private Map<CellAddress, CellStyleAttributes> parseSheet(InputStream sheetStream, StylesTable stylesTable) {
        SheetStyleHandler handler = new SheetStyleHandler(stylesTable, rowLimit);
        try {
            SAXParserFactory factory = SAXParserFactory.newInstance();
            factory.setNamespaceAware(true);
            XMLReader xmlReader = factory.newSAXParser().getXMLReader();
            xmlReader.setContentHandler(handler);
            xmlReader.parse(new InputSource(sheetStream));
        } catch (SAXException e) {
            if (!handler.isLimitReached()) {
                throw new ExcelExportException("Failed to parse sheet XML for styles", e);
            }
            // нормальное завершение по достижению лимита строк
        } catch (Exception e) {
            throw new ExcelExportException("Failed to parse sheet XML for styles", e);
        }
        return handler.getResult();
    }

    // -------------------------------------------------------------------------

    private static class SheetStyleHandler extends DefaultHandler {

        private final StylesTable stylesTable;
        private final int rowLimit;
        private final Map<CellAddress, CellStyleAttributes> result = new LinkedHashMap<>();

        private int currentRow = 0;
        private boolean limitReached = false;

        SheetStyleHandler(StylesTable stylesTable, int rowLimit) {
            this.stylesTable = stylesTable;
            this.rowLimit = rowLimit;
        }

        @Override
        public void startElement(String uri, String localName, String qName, Attributes atts) throws SAXException {
            String tag = localName.isEmpty() ? qName : localName;

            if ("row".equals(tag)) {
                String r = atts.getValue("r");
                currentRow = (r != null) ? Integer.parseInt(r) : currentRow + 1;
                if (currentRow > rowLimit) {
                    limitReached = true;
                    throw new SAXException("row-limit-reached");
                }
                return;
            }

            if ("c".equals(tag)) {
                String cellRef  = atts.getValue("r");
                String styleIdx = atts.getValue("s");
                if (cellRef == null || styleIdx == null) return;
                try {
                    XSSFCellStyle style = stylesTable.getStyleAt(Integer.parseInt(styleIdx));
                    result.put(parseCellRef(cellRef), convertStyle(style));
                } catch (NumberFormatException ignored) {
                    // некорректный индекс стиля — пропускаем
                }
            }
        }

        boolean isLimitReached() { return limitReached; }

        Map<CellAddress, CellStyleAttributes> getResult() { return result; }

        // ----- cell reference parsing ----------------------------------------

        /**
         * Конвертирует строковую ссылку на ячейку (например, {@code "B5"}) в 0-based {@link CellAddress}.
         *
         * @param ref ссылка вида "A1", "BC42" и т.д.
         * @return {@link CellAddress} с 0-based row и col
         */
        private static CellAddress parseCellRef(String ref) {
            int i = 0;
            while (i < ref.length() && Character.isLetter(ref.charAt(i))) i++;
            int col = 0;
            for (char c : ref.substring(0, i).toCharArray()) {
                col = col * 26 + (Character.toUpperCase(c) - 'A' + 1);
            }
            int row = Integer.parseInt(ref.substring(i)) - 1; // 0-based
            return new CellAddress(row, col - 1);              // 0-based
        }

        // ----- style conversion ----------------------------------------------

        /**
         * Конвертирует {@link XSSFCellStyle} POI в независимый DTO {@link CellStyleAttributes}.
         *
         * @param s стиль ячейки POI
         * @return DTO с атрибутами шрифта, заливки, границ, формата и выравнивания
         */
        private static CellStyleAttributes convertStyle(XSSFCellStyle s) {
            XSSFFont font = s.getFont();
            return new CellStyleAttributes()
                    .setBold(font.getBold())
                    .setItalic(font.getItalic())
                    .setUnderline(font.getUnderline() != org.apache.poi.ss.usermodel.FontUnderline.NONE.getByteValue())
                    .setWrapText(s.getWrapText())
                    .setFontName(font.getFontName())
                    .setFontSize(font.getFontHeightInPoints())
                    .setFontColor(colorHex(font.getXSSFColor()))
                    .setFillColor(colorHex(s.getFillForegroundColorColor()))
                    .setNumberFormat(s.getDataFormatString())
                    .setHorizontalAlignment(mapAlignment(s.getAlignment()))
                    .setBorderTopStyle(s.getBorderTop().name())
                    .setBorderBottomStyle(s.getBorderBottom().name())
                    .setBorderLeftStyle(s.getBorderLeft().name())
                    .setBorderRightStyle(s.getBorderRight().name());
        }

        /**
         * Возвращает цвет в формате {@code "RRGGBB"} или {@code null}, если цвет не задан.
         * Сначала пробует {@code getRGBWithTint()}, затем {@code getARGBHex()}.
         *
         * @param color цвет POI или {@code null}
         * @return строка вида {@code "FF0000"} или {@code null}
         */
        private static String colorHex(XSSFColor color) {
            if (color == null) return null;
            byte[] rgb = color.getRGBWithTint();
            if (rgb != null && rgb.length >= 3) {
                return String.format("%02X%02X%02X", rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF);
            }
            String argb = color.getARGBHex();
            if (argb != null && argb.length() >= 6) {
                return argb.substring(argb.length() - 6);
            }
            return null;
        }

        /**
         * Конвертирует POI {@link HorizontalAlignment} в строку, совместимую с fastexcel.
         * Значения {@code GENERAL} и неизвестные возвращают {@code null}.
         *
         * @param a выравнивание POI или {@code null}
         * @return строка выравнивания fastexcel или {@code null}
         */
        private static String mapAlignment(HorizontalAlignment a) {
            if (a == null) return null;
            return switch (a) {
                case LEFT              -> "left";
                case CENTER            -> "center";
                case RIGHT             -> "right";
                case FILL              -> "fill";
                case JUSTIFY           -> "justify";
                case CENTER_SELECTION  -> "centerContinuous";
                case DISTRIBUTED       -> "distributed";
                default                -> null;
            };
        }
    }
}
