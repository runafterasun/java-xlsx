package ru.objectsfill;

import org.dhatim.fastexcel.reader.*;
import ru.objectsfill.enums.BindingMode;
import ru.objectsfill.enums.LoopOperationType;
import ru.objectsfill.exception.ExcelImportException;
import ru.objectsfill.model.ExcelImportParamCore;
import ru.objectsfill.model.FieldParam;
import ru.objectsfill.reader.ExcelTemplateReader;
import ru.objectsfill.reader.TemplateReader;
import ru.objectsfill.util.FieldNameUtils;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;


public class ExcelImportUtil {

    private ExcelImportUtil() {
    }

    /**
     * Читает шаблонный Excel-файл для получения разметки и наполняет объекты данными
     * из import-файла. Делегирует в {@link #importExcel(ExcelImportParamCore, TemplateReader, InputStream)}
     * через {@link ExcelTemplateReader}.
     *
     * @param importParam параметры импорта: набор объектов для наполнения
     * @param tmplXLSX    шаблон, откуда читается разметка (расположение маркеров)
     * @param importXLSX  файл с данными, из которого ведётся чтение по разметке
     * @throws ExcelImportException если произошла ошибка при чтении или наполнении объектов
     */
    public static void importExcel(ExcelImportParamCore importParam, InputStream tmplXLSX, InputStream importXLSX) {
        importExcel(importParam, new ExcelTemplateReader(tmplXLSX), importXLSX);
    }

    /**
     * Наполняет объекты данными из import-файла, используя произвольный источник разметки.
     * Это основная точка входа: позволяет передавать шаблон как Excel ({@link ExcelTemplateReader})
     * или как JSON ({@link ru.objectsfill.reader.JsonTemplateReader}).
     *
     * <p>После чтения маркеров проверяет, что все loop-записи с режимом {@link BindingMode#HEADER}
     * имеют непустой {@code headerName}.
     *
     * @param importParam    параметры импорта: набор объектов для наполнения
     * @param templateReader источник разметки шаблона
     * @param importXLSX     файл с данными, из которого ведётся чтение по разметке
     * @throws ExcelImportException если произошла ошибка чтения, валидации или наполнения объектов
     */
    public static void importExcel(ExcelImportParamCore importParam, TemplateReader templateReader, InputStream importXLSX) {
        List<FieldParam> paramMap = new ArrayList<>();
        templateReader.read(importParam, paramMap);

        Map<String, List<FieldParam>> collectWithParams = paramMap.stream()
                .collect(Collectors.groupingBy(ss -> FieldNameUtils.getKey(ss.getFieldName())));

        importParam.getParamsMap()
                .forEach((key, info) -> {
                    if (collectWithParams.containsKey(key)) {
                        Map<String, List<FieldParam>> groupBySheet = collectWithParams.get(key).stream()
                                .collect(Collectors.groupingBy(FieldParam::getSheetName));
                        info.setParamMap(groupBySheet);
                    }
                });

        validateHeaderBindingHasHeaderName(importParam);

        try (var importWB = new ReadableWorkbook(importXLSX, new ReadingOptions(true, false))) {
            SingleObjectFiller.singleObjectFill(importWB, importParam);
            LoopObjectFiller.loopObjectGetMarkersFromImportFile(importWB, importParam);
            LoopObjectFiller.loopObjectFill(importWB, importParam);
        } catch (ExcelImportException e) {
            throw e;
        } catch (Exception e) {
            throw new ExcelImportException("Failed to import Excel file", e);
        }
    }

    /**
     * Проверяет, что каждый loop-маркер с режимом {@link BindingMode#HEADER} имеет непустой
     * {@code headerName}. Вызывается после того, как {@link TemplateReader} заполнил {@code paramMap}.
     *
     * @param importParam параметры импорта
     * @throws ExcelImportException если найден маркер без {@code headerName} при режиме {@code HEADER}
     */
    private static void validateHeaderBindingHasHeaderName(ExcelImportParamCore importParam) {
        importParam.getParamsMap().forEach((key, info) -> {
            if (info.getBindingMode() != BindingMode.HEADER || info.getParamMap().isEmpty()) return;
            info.getParamMap().values().stream()
                    .flatMap(List::stream)
                    .filter(fp -> fp.getFieldName().startsWith(LoopOperationType.FOR.getType()))
                    .filter(fp -> fp.getHeaderName() == null || fp.getHeaderName().isBlank())
                    .findFirst()
                    .ifPresent(fp -> {
                        throw new ExcelImportException(
                                "FieldParam '" + fp.getFieldName() + "' has HEADER binding but 'headerName' is empty");
                    });
        });
    }
}
