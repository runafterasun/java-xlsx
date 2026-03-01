package ru.objectsfill;

import org.dhatim.fastexcel.reader.*;

import java.io.InputStream;
import java.util.*;
import java.util.function.Predicate;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public class ExcelImportService {

    private ExcelImportService(){}

    public static <T> void importExcel(ExcelImportParam<T> importParam, InputStream tmplXLSX, InputStream importXLSX) {
        Map<String, CellAddress> valMap = new HashMap<>();

        try (var tmlWB = new ReadableWorkbook(tmplXLSX, new ReadingOptions(true, false));
             var importWB = new ReadableWorkbook(importXLSX, new ReadingOptions(true, false));) {

            fillParamMaps(tmlWB, importParam.getParamsMap(), valMap, importParam.getWbSheetNumber());

            var importSheet = importWB.getSheet(importParam.getWbSheetNumber());
            if (importSheet.isPresent()) {

                var notLoopObject = getObjectForFill(importParam.getParamsMap(), lfFillObj -> !lfFillObj.getKey().contains(LoopOperationType.FOR.getType()));
                notLoopObject.ifPresent(object -> singleObjectFill(importSheet.get(), valMap, object));

                var loopObject = getObjectForFill(importParam.getParamsMap(), lfFillObj -> lfFillObj.getKey().contains(LoopOperationType.FOR.getType()));
                loopObject.ifPresent(object -> loopFieldRead(importSheet.get(), valMap, importParam.getLoopLst(), object));

            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    public static void fillParamMaps(ReadableWorkbook tmlWB, Map<String, Object> paramsMap, Map<String, CellAddress> valMap, Integer wbSheetNumber) {
        var sheet = tmlWB.getSheet(wbSheetNumber);
        if (sheet.isPresent()) {
            try (Stream<Row> rows = sheet.get().openStream()) {
                rows.forEach(cells -> cells.stream()
                        .filter(Objects::nonNull)
                        .forEach(cell -> {
                            if (cell.getText().lastIndexOf(".") != -1
                                    && validateParam(cell, paramsMap)) {
                                valMap.put(cell.getText(), cell.getAddress());
                            }
                            if (cell.getText().equals(LoopOperationType.FOR.getType())) {
                                valMap.put(LoopOperationType.FOR.getType(), cell.getAddress());
                            }

                        }));
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
        }
    }

    private static void singleObjectFill(Sheet importSheet, Map<String, CellAddress> valMap, Object loopObject) {
        try (var importRow = importSheet.openStream()) {
            importRow.limit(valMap.get(LoopOperationType.FOR.getType()).getRow())
                    .forEach(cells -> cells
                            .stream()
                            .filter(Objects::nonNull)
                            .forEach(cell -> {
                                if (valMap.containsValue(cell.getAddress())) {
                                    keys(valMap, cell.getAddress()).ifPresent(key -> {
                                        if (key.lastIndexOf(".") != -1) {
                                            validateAndCheckCell(cell, key, loopObject);
                                        }
                                    });
                                }
                            }));
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }


    @SuppressWarnings("unchecked")
    private static <T> void loopFieldRead(Sheet importSheet, Map<String, CellAddress> valMap, List<T> loopLst, Object loopObject) {
        try (var importRow = importSheet.openStream()) {
            Map<String, Integer> loopMap = convert(valMap);
            importRow.skip(valMap.get(LoopOperationType.FOR.getType()).getRow() + 1).forEach(cells -> {
                    try {
                        T obj = (T) loopObject.getClass().getConstructor().newInstance();
                        cells.stream().filter(Objects::nonNull)
                                .forEach(cell -> {
                                    if (loopMap.containsValue(cell.getAddress().getColumn())) {
                                        keys(loopMap, cell.getAddress().getColumn()).ifPresent(key -> {
                                            if (key.lastIndexOf(".") != -1) {
                                                validateAndCheckCell(cell, key, obj);
                                            }
                                        });
                                    }
                                });
                        loopLst.add(obj);
                    } catch (Exception e) {
                        throw new RuntimeException(e);
                    }
            });
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    private static boolean validateParam(Cell cell, Map<String, Object> paramsMap) {
        return paramsMap.containsKey(cell.getText().substring(0, cell.getText().lastIndexOf(".")));
    }

    private static <T> void validateAndCheckCell(Cell cell, String key, T object) {
        CellTypeOperation.findByType(cell.getType().name()).validateAndWrite(cell, key, object);
    }

    private static <K, V> Optional<K> keys(Map<K, V> map, V value) {
        return map.entrySet()
                .stream()
                .filter(entry -> value.equals(entry.getValue()))
                .map(Map.Entry::getKey)
                .findFirst();
    }

    private static Map<String, Integer> convert(Map<String, CellAddress> map) {
        return map.entrySet()
                .stream()
                .filter(key -> key.getKey().contains(LoopOperationType.FOR.getType()))
                .filter(key -> !key.getKey().equals(LoopOperationType.FOR.getType()))
                .collect(Collectors.toMap(Map.Entry::getKey, value -> value.getValue().getColumn()));
    }

    private static Optional<Object> getObjectForFill(Map<String, Object> paramsMap, Predicate<Map.Entry<String, Object>> getLoopOrNot) {
        return paramsMap.entrySet()
                .stream()
                .filter(getLoopOrNot)
                .findAny()
                .map(Map.Entry::getValue);
    }
}
