package ru.objectsfill.reader;

import org.dhatim.fastexcel.reader.*;
import ru.objectsfill.enums.LoopOperationType;
import ru.objectsfill.exception.ExcelImportException;
import ru.objectsfill.model.ExcelImportParamCore;
import ru.objectsfill.model.FieldParam;
import ru.objectsfill.model.ImportInformation;

import java.io.InputStream;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * Читает параметры шаблона из Excel-файла.
 * Сканирует каждый лист и распознаёт ячейки с маркерами вида {@code "ключ.поле"}.
 * Для loop-маркеров ({@code "for.*"}) дополнительно читает строку заголовков
 * (строка выше маркера) и сохраняет текст заголовка в {@link FieldParam#setHeaderName}.
 */
public class ExcelTemplateReader implements TemplateReader {

    private final InputStream templateStream;

    /**
     * @param templateStream поток Excel-шаблона с маркерами
     */
    public ExcelTemplateReader(InputStream templateStream) {
        this.templateStream = templateStream;
    }

    /**
     * Читает маркеры из Excel-шаблона и добавляет их в {@code out}.
     * Использует однопроходное чтение: заголовок берётся из предыдущей строки
     * без повторного открытия потока.
     *
     * @param importParam параметры импорта с зарегистрированными объектами
     * @param out         список, в который добавляются найденные {@link FieldParam}
     * @throws ExcelImportException если не удалось прочитать шаблон
     */
    @Override
    public void read(ExcelImportParamCore importParam, List<FieldParam> out) {
        try (var workbook = new ReadableWorkbook(templateStream, new ReadingOptions(true, false))) {
            workbook.getSheets().forEach(sheet -> {
                try (var rows = sheet.openStream()) {
                    Row[] prevRow = {null};
                    rows.forEach(row -> {
                        row.stream()
                                .filter(Objects::nonNull)
                                .forEach(cell -> {
                                    if (isKnownMarker(cell, importParam.getParamsMap())) {
                                        FieldParam fp = new FieldParam()
                                                .setFieldName(cell.getText().trim())
                                                .setCellAddress(cell.getAddress())
                                                .setSheetName(sheet.getName());
                                        findHeader(fp, prevRow[0]);
                                        out.add(fp);
                                    }
                                });
                        prevRow[0] = row;
                    });
                } catch (Exception e) {
                    throw new ExcelImportException("Failed to process template sheet: " + sheet.getName(), e);
                }
            });
        } catch (ExcelImportException e) {
            throw e;
        } catch (Exception e) {
            throw new ExcelImportException("Failed to read Excel template", e);
        }
    }

    /**
     * Для loop-маркера ищет текст заголовка в строке, непосредственно предшествующей
     * строке маркера, по тому же столбцу.
     *
     * @param fp      параметр, для которого ищем заголовок
     * @param prevRow предыдущая строка листа (строка выше маркера), или {@code null}
     */
    private static void findHeader(FieldParam fp, Row prevRow) {
        if (fp.getFieldName().startsWith(LoopOperationType.FOR.getType()) && prevRow != null) {
            prevRow.stream()
                    .filter(Objects::nonNull)
                    .filter(hc -> hc.getColumnIndex() == fp.getCellAddress().getColumn())
                    .findFirst()
                    .ifPresent(hc -> fp.setHeaderName(hc.getText().trim()));
        }
    }

    /**
     * Проверяет, является ли текст ячейки маркером одного из зарегистрированных объектов.
     * Маркер должен содержать точку, а часть до первой точки — совпадать с ключом в {@code paramsMap}.
     *
     * @param cell      ячейка шаблона
     * @param paramsMap карта зарегистрированных объектов
     * @return {@code true}, если ячейка содержит известный маркер
     */
    private static boolean isKnownMarker(Cell cell, Map<String, ImportInformation> paramsMap) {
        String text = cell.getText();
        int dot = text.lastIndexOf(".");
        if (dot == -1) return false;
        return paramsMap.containsKey(text.substring(0, dot));
    }
}
