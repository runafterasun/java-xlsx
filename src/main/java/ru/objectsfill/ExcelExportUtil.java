package ru.objectsfill;

import org.dhatim.fastexcel.Workbook;
import org.dhatim.fastexcel.Worksheet;
import org.dhatim.fastexcel.reader.CellAddress;
import ru.objectsfill.enums.LoopOperationType;
import ru.objectsfill.exception.ExcelExportException;
import ru.objectsfill.model.ExcelExportParamCore;
import ru.objectsfill.model.ExcelImportParamCore;
import ru.objectsfill.model.ExportInformation;
import ru.objectsfill.model.FieldParam;
import ru.objectsfill.model.ImportInformation;
import ru.objectsfill.reader.ExcelTemplateReader;
import ru.objectsfill.reader.TemplateReader;
import ru.objectsfill.style.CellStyleAttributes;
import ru.objectsfill.style.PoiStyleReader;
import ru.objectsfill.util.FieldNameUtils;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

public class ExcelExportUtil {

    private ExcelExportUtil() {}

    /**
     * Экспортирует данные в Excel-файл, используя Excel-шаблон для разметки.
     *
     * @param exportParam параметры экспорта с данными
     * @param tmplXLSX    шаблон с маркерами
     * @param output      поток вывода для результирующего файла
     */
    public static void exportExcel(ExcelExportParamCore exportParam, InputStream tmplXLSX, OutputStream output) {
        exportExcel(exportParam, new ExcelTemplateReader(tmplXLSX), output);
    }

    /**
     * Экспортирует данные в Excel-файл, используя произвольный источник разметки.
     *
     * @param exportParam    параметры экспорта с данными
     * @param templateReader источник разметки (Excel или JSON)
     * @param output         поток вывода для результирующего файла
     */
    public static void exportExcel(ExcelExportParamCore exportParam, TemplateReader templateReader, OutputStream output) {
        List<FieldParam> allParams = readTemplate(templateReader, exportParam);

        distributeParams(allParams, exportParam);

        List<String> sheetNames = allParams.stream()
                .map(FieldParam::getSheetName)
                .distinct()
                .collect(Collectors.toList());

        try (Workbook wb = new Workbook(output, "excel-export", "1.0")) {
            for (String sheetName : sheetNames) {
                Worksheet ws = wb.newWorksheet(sheetName);
                writeSingleObjects(ws, sheetName, exportParam);
                writeLoopObjects(ws, sheetName, exportParam);
            }
        } catch (ExcelExportException e) {
            throw e;
        } catch (Exception e) {
            throw new ExcelExportException("Failed to export Excel", e);
        }
    }

    /**
     * Экспортирует данные с переносом стилей из шаблона.
     * Apache POI читает стили потоково (первые {@link PoiStyleReader#DEFAULT_STYLE_READ_DEPTH} строк).
     * Стили применяются только к ячейкам одиночных объектов; табличные строки пишутся без стилей.
     *
     * @param exportParam параметры экспорта с данными
     * @param tmplXLSX    шаблон — источник разметки и стилей
     * @param output      поток вывода для результирующего файла
     */
    public static void exportExcelWithStyles(ExcelExportParamCore exportParam,
                                             InputStream tmplXLSX, OutputStream output) {
        exportExcelWithStyles(exportParam, tmplXLSX, output, PoiStyleReader.DEFAULT_STYLE_READ_DEPTH);
    }

    /**
     * Экспортирует данные с переносом стилей, читая стили из первых {@code styleReadDepth} строк шаблона.
     *
     * @param exportParam    параметры экспорта с данными
     * @param tmplXLSX       шаблон — источник разметки и стилей
     * @param output         поток вывода для результирующего файла
     * @param styleReadDepth количество строк шаблона, из которых читаются стили
     */
    public static void exportExcelWithStyles(ExcelExportParamCore exportParam,
                                             InputStream tmplXLSX, OutputStream output,
                                             int styleReadDepth) {
        byte[] templateBytes;
        try {
            templateBytes = tmplXLSX.readAllBytes();
        } catch (IOException e) {
            throw new ExcelExportException("Failed to read template stream", e);
        }

        Map<String, Map<CellAddress, CellStyleAttributes>> stylesBySheet =
                new PoiStyleReader(styleReadDepth).read(new ByteArrayInputStream(templateBytes));

        exportWithStyles(exportParam,
                new ExcelTemplateReader(new ByteArrayInputStream(templateBytes)),
                output,
                stylesBySheet);
    }

    /**
     * Выполняет экспорт с применением переданных стилей к ячейкам одиночных объектов.
     *
     * @param exportParam    параметры экспорта с данными
     * @param templateReader источник разметки (Excel или JSON)
     * @param output         поток вывода для результирующего файла
     * @param stylesBySheet  стили ячеек по имени листа, прочитанные из шаблона
     */
    private static void exportWithStyles(ExcelExportParamCore exportParam, TemplateReader templateReader,
                                         OutputStream output,
                                         Map<String, Map<CellAddress, CellStyleAttributes>> stylesBySheet) {
        List<FieldParam> allParams = readTemplate(templateReader, exportParam);
        distributeParams(allParams, exportParam);

        List<String> sheetNames = allParams.stream()
                .map(FieldParam::getSheetName)
                .distinct()
                .collect(Collectors.toList());

        try (Workbook wb = new Workbook(output, "excel-export", "1.0")) {
            for (String sheetName : sheetNames) {
                Worksheet ws = wb.newWorksheet(sheetName);
                Map<CellAddress, CellStyleAttributes> sheetStyles =
                        stylesBySheet.getOrDefault(sheetName, Map.of());
                writeSingleObjectsWithStyles(ws, sheetName, exportParam, sheetStyles);
                writeLoopObjects(ws, sheetName, exportParam);
            }
        } catch (ExcelExportException e) {
            throw e;
        } catch (Exception e) {
            throw new ExcelExportException("Failed to export Excel with styles", e);
        }
    }

    /**
     * Записывает ячейки одиночных объектов на лист с применением стилей из шаблона.
     *
     * @param ws          лист для записи
     * @param sheetName   имя листа (используется как ключ поиска параметров)
     * @param exportParam параметры экспорта
     * @param sheetStyles стили ячеек для данного листа
     */
    private static void writeSingleObjectsWithStyles(Worksheet ws, String sheetName,
                                                     ExcelExportParamCore exportParam,
                                                     Map<CellAddress, CellStyleAttributes> sheetStyles) {
        exportParam.getParamsMap().forEach((key, info) -> {
            if (key.startsWith(LoopOperationType.FOR.getType())) return;
            if (info.getDataList().isEmpty()) return;
            List<FieldParam> params = info.getParamMap().get(sheetName);
            if (params == null) return;
            SingleObjectWriter.write(ws, params, info.getDataList().get(0), sheetStyles, exportParam.getWarnings());
        });
    }

    // -------------------------------------------------------------------------

    /**
     * Читает разметку из шаблона, создавая временный {@link ExcelImportParamCore} с ключами из {@code exportParam}.
     *
     * @param reader      источник разметки
     * @param exportParam параметры экспорта (ключи используются для регистрации в dummy-импорте)
     * @return список всех {@link FieldParam}, прочитанных из шаблона
     */
    private static List<FieldParam> readTemplate(TemplateReader reader, ExcelExportParamCore exportParam) {
        ExcelImportParamCore dummy = new ExcelImportParamCore();
        exportParam.getParamsMap().forEach((key, info) ->
                dummy.getParamsMap().put(key, new ImportInformation().setClazz(Object.class)));

        List<FieldParam> params = new ArrayList<>();
        reader.read(dummy, params);
        return params;
    }

    /**
     * Раскладывает прочитанные {@link FieldParam} по {@link ExportInformation#setParamMap(Map)}.
     * Группирует сначала по ключу маркера, затем по имени листа.
     *
     * @param allParams   список параметров, прочитанных из шаблона
     * @param exportParam параметры экспорта, куда записываются сгруппированные параметры
     */
    private static void distributeParams(List<FieldParam> allParams, ExcelExportParamCore exportParam) {
        Map<String, List<FieldParam>> byKey = allParams.stream()
                .collect(Collectors.groupingBy(fp -> FieldNameUtils.getKey(fp.getFieldName())));

        byKey.forEach((key, fps) -> {
            ExportInformation info = exportParam.getParamsMap().get(key);
            if (info == null) return;
            Map<String, List<FieldParam>> bySheet = fps.stream()
                    .collect(Collectors.groupingBy(FieldParam::getSheetName));
            info.setParamMap(bySheet);
        });
    }

    /**
     * Записывает ячейки одиночных объектов на лист (без стилей).
     *
     * @param ws          лист для записи
     * @param sheetName   имя листа
     * @param exportParam параметры экспорта
     */
    private static void writeSingleObjects(Worksheet ws, String sheetName, ExcelExportParamCore exportParam) {
        exportParam.getParamsMap().forEach((key, info) -> {
            if (key.startsWith(LoopOperationType.FOR.getType())) return;
            if (info.getDataList().isEmpty()) return;
            List<FieldParam> params = info.getParamMap().get(sheetName);
            if (params == null) return;
            SingleObjectWriter.write(ws, params, info.getDataList().get(0), exportParam.getWarnings());
        });
    }

    /**
     * Записывает строки loop-блоков на лист.
     * Для каждого {@code for.}-ключа заголовки пишутся на строку выше маркера,
     * данные — начиная со строки маркера.
     *
     * @param ws          лист для записи
     * @param sheetName   имя листа
     * @param exportParam параметры экспорта
     */
    private static void writeLoopObjects(Worksheet ws, String sheetName, ExcelExportParamCore exportParam) {
        exportParam.getParamsMap().forEach((key, info) -> {
            if (!key.startsWith(LoopOperationType.FOR.getType())) return;
            if (info.getDataList().isEmpty()) return;
            List<FieldParam> params = info.getParamMap().get(sheetName);
            if (params == null) return;
            LoopObjectWriter.write(ws, params, info.getDataList(), exportParam.getWarnings());
        });
    }
}
