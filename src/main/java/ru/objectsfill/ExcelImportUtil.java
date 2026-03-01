package ru.objectsfill;

import org.dhatim.fastexcel.reader.*;
import ru.objectsfill.enums.CellTypeOperation;
import ru.objectsfill.enums.LoopOperationType;
import ru.objectsfill.exception.ExcelImportException;
import ru.objectsfill.model.ExcelImportParamCore;
import ru.objectsfill.model.FieldParam;
import ru.objectsfill.model.ImportInformation;

import java.io.IOException;
import java.io.InputStream;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.function.Function;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import static ru.objectsfill.enums.CellTypeOperation.getSubstringKey;
import static ru.objectsfill.util.FileUtils.fillObject;


public class ExcelImportUtil {

    private ExcelImportUtil() {
    }

    /**
     * Читает шаблонный файл для получения разметки и наполняет объекты данными из import-файла.
     *
     * @param importParam параметры импорта: набор объектов для наполнения и номер вкладки
     * @param tmplXLSX    шаблон, откуда читается разметка (расположение маркеров)
     * @param importXLSX  файл с данными, из которого ведётся чтение по разметке
     * @throws ExcelImportException если произошла ошибка при чтении или наполнении объектов
     */
    public static void importExcel(ExcelImportParamCore importParam, InputStream tmplXLSX, InputStream importXLSX) {
        List<FieldParam> paramMap = new ArrayList<>();

        try (var tmlWB = new ReadableWorkbook(tmplXLSX, new ReadingOptions(true, false));
             var importWB = new ReadableWorkbook(importXLSX, new ReadingOptions(true, false))) {

            fillParamMaps(tmlWB, importParam, paramMap);

            Map<String, List<FieldParam>> collectWithParams = paramMap
                    .stream()
                    .collect(Collectors.groupingBy(ss -> substringParam(ss.getFieldName())));

            importParam.getParamsMap()
                    .forEach((objIndet, objData) -> {
                        if (collectWithParams.containsKey(objIndet)) {
                            Map<String, List<FieldParam>> groupBySheet = collectWithParams.get(objIndet)
                                    .stream()
                                    .collect(Collectors.groupingBy(FieldParam::getSheetName));
                            objData.setParamMap(groupBySheet);
                        }
                    });

            singleObjectFill(importWB, importParam);
            loopObjectGetMarkersFromImportFile(importWB, importParam);
            loopObjectFill(importWB, importParam);

        } catch (ExcelImportException e) {
            throw e;
        } catch (Exception e) {
            throw new ExcelImportException("Failed to import Excel file", e);
        }
    }

    /**
     * Сканирует листы шаблонной книги и собирает все найденные маркеры полей в список paramMap.
     *
     * @param tmlWB       книга-шаблон, откуда читается разметка
     * @param importParam параметры импорта с картой ожидаемых объектов
     * @param paramMap    список, в который добавляются найденные {@link FieldParam}
     * @throws ExcelImportException если не удалось прочитать строки листа
     */
    public static void fillParamMaps(ReadableWorkbook tmlWB, ExcelImportParamCore importParam, List<FieldParam> paramMap) {
        tmlWB.getSheets().forEach(sheet -> {
            try (Stream<Row> rows = sheet.openStream()) {
                rows.forEach(cells -> cells.stream()
                        .filter(Objects::nonNull)
                        .forEach(cell -> {
                            if (validateParam(cell, importParam.getParamsMap())) {
                                FieldParam map = new FieldParam().setFieldName(cell.getText().trim())
                                        .setCellAddress(cell.getAddress())
                                        .setSheetName(sheet.getName());
                                findHeaders(cells, map, sheet, cell);
                                paramMap.add(map);
                            }
                        }));
            } catch (Exception e) {
                throw new ExcelImportException("Failed to process template sheet: " + sheet.getName(), e);
            }
        });
    }

    /**
     * Для loop-маркера ищет строку заголовков (строку выше маркера) и записывает
     * название заголовка того же столбца в {@code paramMap}.
     *
     * @param row      строка, содержащая маркер
     * @param paramMap параметр, для которого ищем заголовок
     * @param sheet    лист, на котором расположен маркер
     * @param cell     ячейка с маркером
     * @throws ExcelImportException если не удалось открыть поток строк листа
     */
    private static void findHeaders(Row row, FieldParam paramMap, Sheet sheet, Cell cell) {
        if (paramMap.getFieldName().startsWith(LoopOperationType.FOR.getType())) {
            try {
                sheet.openStream()
                        .filter(headerRow -> headerRow.getRowNum() == (row.getRowNum() - 1))
                        .findFirst()
                        .flatMap(headerCells -> headerCells
                                .stream()
                                .filter(Objects::nonNull)
                                .filter(headerCell -> headerCell.getColumnIndex() == cell.getColumnIndex())
                                .findFirst())
                        .ifPresent(headerCell -> paramMap.setHeaderName(headerCell.getText().trim()));
            } catch (IOException e) {
                throw new ExcelImportException("Failed to read header row in sheet: " + sheet.getName(), e);
            }
        }
    }

    /**
     * Наполняет одиночные (не loop) объекты данными из import-файла.
     * Для каждого не-FOR ключа создаёт экземпляр целевого класса, считывает значения
     * ячеек по координатам из шаблона и добавляет объект в {@link ImportInformation#getLoopLst()}.
     *
     * @param readableWB  книга с данными для импорта
     * @param importParam параметры импорта
     * @throws ExcelImportException если не удалось создать экземпляр класса или прочитать ячейки
     */
    private static void singleObjectFill(ReadableWorkbook readableWB, ExcelImportParamCore importParam) {
        readableWB.getSheets()
                .forEach(sheet -> importParam
                        .getParamsMap()
                        .entrySet()
                        .stream()
                        .filter(dd -> !dd.getKey().startsWith(LoopOperationType.FOR.getType()))
                        .forEach(valMap -> {
                            if (valMap.getValue().getParamMap().containsKey(sheet.getName())) {
                                Map<CellAddress, FieldParam> collect = valMap.getValue().getParamMap().get(sheet.getName())
                                        .stream()
                                        .filter(ss -> !ss.getFieldName().startsWith(LoopOperationType.FOR.getType()))
                                        .collect(Collectors.toMap(FieldParam::getCellAddress, Function.identity()));
                                try {
                                    Object object = valMap.getValue().getClazz().getConstructor().newInstance();
                                    sheet.openStream()
                                            .limit(100)
                                            .forEach(ss -> ss.stream()
                                                    .filter(Objects::nonNull)
                                                    .forEach(cell -> {
                                                        if (collect.containsKey(cell.getAddress())) {
                                                            FieldParam fieldParam = collect.get(cell.getAddress());
                                                            validateAndWrite(cell, fieldParam.getFieldName(), object);
                                                        }
                                                    }));
                                    valMap.getValue().getLoopLst().add(object);
                                } catch (Exception e) {
                                    throw new ExcelImportException(
                                            "Failed to instantiate or fill object for key: " + valMap.getKey(), e);
                                }
                            }
                        }));
    }

    /**
     * Читает import-файл и устанавливает {@link FieldParam#setCurrentCellAddress} для каждого
     * loop-маркера, совпадающего с заголовком столбца.
     * Также удаляет найденные поля из списка обязательных.
     *
     * @param readableWB  книга с данными для импорта
     * @param importParam параметры импорта
     * @throws ExcelImportException если не удалось прочитать строки листа
     */
    private static void loopObjectGetMarkersFromImportFile(ReadableWorkbook readableWB, ExcelImportParamCore importParam) {
        readableWB.getSheets()
                .forEach(sheet -> {
                    try {
                        sheet.openStream()
                                .limit(100)
                                .forEach(row -> row
                                        .stream()
                                        .filter(Objects::nonNull)
                                        .forEach(cell -> importParam
                                                .getParamsMap()
                                                .entrySet()
                                                .stream()
                                                .filter(entry -> entry.getKey().startsWith(LoopOperationType.FOR.getType()))
                                                .forEach(valMap -> valMap.getValue().getParamMap()
                                                        .entrySet()
                                                        .stream()
                                                        .filter(entry -> entry.getKey().equals(sheet.getName()))
                                                        .forEach(entry -> entry.getValue()
                                                                .stream()
                                                                .filter(fieldParam -> fieldParam.getHeaderName().trim().equals(cell.getRawValue()))
                                                                .forEach(fieldParam -> {
                                                                    fieldParam.setCurrentCellAddress(cell.getAddress());
                                                                    valMap.getValue().getRequiredFields().remove(getParamName(fieldParam.getFieldName()));
                                                                }))
                                                ))
                                );
                    } catch (IOException e) {
                        throw new ExcelImportException("Failed to read import file sheet: " + sheet.getName(), e);
                    }
                });
    }

    /**
     * Для каждого loop-ключа читает строки import-файла (начиная после строки заголовков),
     * создаёт экземпляр целевого класса для каждой непустой строки и наполняет его данными.
     *
     * @param readableWB  книга с данными для импорта
     * @param importParam параметры импорта
     * @throws ExcelImportException если не удалось прочитать строки или создать экземпляр объекта
     */
    private static void loopObjectFill(ReadableWorkbook readableWB, ExcelImportParamCore importParam) {
        readableWB.getSheets().forEach(sheet -> {

            Optional<Map.Entry<String, ImportInformation>> first = getStringImportInformationEntry(importParam, sheet);

            if (first.isPresent()) {
                List<FieldParam> paramMaps = first.get().getValue().getParamMap().get(sheet.getName());
                int skipRow = getSkipRow(paramMaps);
                AtomicInteger counter = new AtomicInteger(skipRow);

                try {
                    sheet.openStream()
                            .filter(cells -> cells.getRowNum() > skipRow)
                            .forEach(cells -> {
                                try {
                                    long count = getCount(cells);
                                    if (cells.getRowNum() - counter.get() < 2 && count > 1L) {
                                        Object fillObj = first.get().getValue().getClazz().getConstructor().newInstance();
                                        fillObjectLoop(cells, paramMaps, fillObj);
                                        first.get().getValue().getLoopLst().add(fillObj);
                                        counter.set(cells.getRowNum());
                                    }
                                } catch (Exception e) {
                                    throw new ExcelImportException("Failed to instantiate loop object", e);
                                }
                            });
                } catch (ExcelImportException e) {
                    throw e;
                } catch (Exception e) {
                    throw new ExcelImportException("Failed to process loop rows in sheet: " + sheet.getName(), e);
                }
            }
        });
    }

    /**
     * Подсчитывает количество непустых ячеек в строке.
     *
     * @param cells строка листа
     * @return количество ячеек с типом, отличным от {@link CellType#EMPTY}
     */
    private static long getCount(Row cells) {
        return cells.stream()
                .filter(Objects::nonNull)
                .filter(cell -> cell.getType() != CellType.EMPTY)
                .count();
    }

    /**
     * Возвращает номер строки, с которой начинаются данные (строка после заголовков).
     *
     * @param paramMaps список маркеров с установленными адресами заголовков
     * @return номер строки заголовка + 1, или 1, если маркеры не найдены
     */
    private static int getSkipRow(List<FieldParam> paramMaps) {
        return paramMaps
                .stream()
                .map(fieldParam -> fieldParam.getCurrentCellAddress().getRow())
                .findAny()
                .orElse(0) + 1;
    }

    /**
     * Находит первый loop-ключ в {@code importParam}, у которого есть данные для текущего листа.
     *
     * @param importParam параметры импорта
     * @param sheet       текущий лист
     * @return {@link Optional} с записью ключ-значение, или пустой Optional, если совпадений нет
     */
    private static Optional<Map.Entry<String, ImportInformation>> getStringImportInformationEntry(
            ExcelImportParamCore importParam, Sheet sheet) {
        return importParam.getParamsMap()
                .entrySet()
                .stream()
                .filter(entry -> entry.getKey().startsWith(LoopOperationType.FOR.getType()))
                .filter(entry -> entry.getValue().getParamMap().containsKey(sheet.getName()))
                .findFirst();
    }

    /**
     * Наполняет единственный объект данными из одной строки loop-таблицы.
     * Для каждой ячейки ищет соответствующий маркер по столбцу и записывает значение.
     *
     * @param cells     строка с данными
     * @param paramMaps список маркеров для текущего loop-блока
     * @param obj       объект, который наполняется
     */
    private static <T> void fillObjectLoop(Row cells, List<FieldParam> paramMaps, T obj) {
        for (int cellNum = 0; cellNum < cells.getCellCount(); cellNum++) {
            if (cells.getCell(cellNum) != null) {
                Cell cell = cells.getCell(cellNum);
                Optional<FieldParam> any = getFieldParam(paramMaps, cell);

                if (any.isPresent()) {
                    String keys = any.get().getFieldName();
                    if (cell.getDataFormatString() != null &&
                            cell.getDataFormatString().contains("m") &&
                            cells.getCellAsDate(cellNum).isPresent()) {
                        var fieldName = getSubstringKey(keys);
                        if (cells.getCellAsDate(cellNum).isPresent()) {
                            String data = cells.getCellAsDate(cellNum).get().toString();
                            fillObject(obj, data, fieldName);
                        }
                    } else {
                        validateAndWrite(cell, keys, obj);
                    }
                }
            }
        }
    }

    /**
     * Ищет маркер, чей столбец совпадает со столбцом ячейки, а адрес заголовка расположен выше.
     *
     * @param paramMaps список маркеров loop-блока
     * @param cell      ячейка с данными
     * @return найденный маркер, или пустой Optional
     */
    private static Optional<FieldParam> getFieldParam(List<FieldParam> paramMaps, Cell cell) {
        return paramMaps
                .stream()
                .filter(fieldParam -> fieldParam.getCurrentCellAddress().getRow() < cell.getAddress().getRow() &&
                        fieldParam.getCurrentCellAddress().getColumn() == cell.getAddress().getColumn())
                .findAny();
    }

    /**
     * Проверяет, является ли текст ячейки маркером одного из ожидаемых объектов импорта.
     * Маркер имеет формат {@code "ключ.поле"}.
     *
     * @param cell      ячейка шаблона
     * @param paramsMap карта ожидаемых объектов импорта
     * @return {@code true}, если ячейка содержит известный маркер
     */
    private static boolean validateParam(Cell cell, Map<String, ImportInformation> paramsMap) {
        if (cell.getText().lastIndexOf(".") != -1) {
            return paramsMap.containsKey(cell.getText().substring(0, cell.getText().lastIndexOf(".")));
        }
        return false;
    }

    /**
     * Извлекает префикс (идентификатор объекта) из полного имени маркера.
     * Например, из {@code "for.dateList.rate"} вернёт {@code "for.dateList"}.
     *
     * @param field полное имя маркера
     * @return часть строки до последней точки, или исходная строка, если точки нет
     */
    private static String substringParam(String field) {
        if (field.lastIndexOf(".") != -1) {
            return field.substring(0, field.lastIndexOf("."));
        }
        return field;
    }

    /**
     * Извлекает имя поля из полного имени маркера (часть после последней точки).
     * Например, из {@code "for.dateList.rate"} вернёт {@code "rate"}.
     *
     * @param field полное имя маркера
     * @return имя поля, или исходная строка, если точки нет
     */
    private static String getParamName(String field) {
        if (field.lastIndexOf(".") != -1) {
            return field.substring(field.lastIndexOf(".") + 1);
        }
        return field;
    }

    /**
     * Определяет тип ячейки и записывает её значение в соответствующее поле объекта.
     *
     * @param cell   ячейка с данными
     * @param key    полное имя маркера для определения целевого поля
     * @param object объект, который наполняется
     */
    private static <T> void validateAndWrite(Cell cell, String key, T object) {
        CellTypeOperation.findByType(cell.getType().name()).validateAndWrite(cell, key, object);
    }
}
