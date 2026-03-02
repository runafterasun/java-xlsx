package ru.objectsfill;

import org.dhatim.fastexcel.reader.*;
import ru.objectsfill.enums.BindingMode;
import ru.objectsfill.enums.CellTypeOperation;
import ru.objectsfill.enums.LoopOperationType;
import ru.objectsfill.exception.ExcelImportException;
import ru.objectsfill.model.ExcelImportParamCore;
import ru.objectsfill.model.FieldParam;
import ru.objectsfill.model.ImportInformation;
import ru.objectsfill.util.FieldNameUtils;
import ru.objectsfill.util.FileUtils;

import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;

class LoopObjectFiller {

    private LoopObjectFiller() {}

    private static final int HEADER_SEARCH_DEPTH = 100;

    static void loopObjectGetMarkersFromImportFile(ReadableWorkbook readableWB, ExcelImportParamCore importParam) {
        applyPositionBinding(importParam);
        applyHeaderBinding(readableWB, importParam);
    }

    private static void applyPositionBinding(ExcelImportParamCore importParam) {
        importParam.getParamsMap()
                .entrySet()
                .stream()
                .filter(entry -> entry.getKey().startsWith(LoopOperationType.FOR.getType()))
                .filter(entry -> entry.getValue().getBindingMode() == BindingMode.POSITION)
                .forEach(entry -> {
                    entry.getValue().getParamMap().values()
                            .forEach(fieldParams -> fieldParams
                                    .forEach(fp -> fp.setCurrentCellAddress(fp.getCellAddress())));
                    entry.getValue().getRequiredFields().clear();
                });
    }

    private static void applyHeaderBinding(ReadableWorkbook readableWB, ExcelImportParamCore importParam) {
        List<Map.Entry<String, ImportInformation>> relevantEntries = importParam.getParamsMap()
                .entrySet()
                .stream()
                .filter(entry -> entry.getKey().startsWith(LoopOperationType.FOR.getType()))
                .filter(entry -> entry.getValue().getBindingMode() == BindingMode.HEADER)
                .filter(entry -> !entry.getValue().getParamMap().isEmpty())
                .collect(Collectors.toList());

        if (relevantEntries.isEmpty()) return;

        readableWB.getSheets().forEach(sheet -> {
            try {
                matchHeadersOnSheet(sheet, relevantEntries);
            } catch (IOException e) {
                throw new ExcelImportException("Failed to read import file sheet: " + sheet.getName(), e);
            }
        });
    }

    private static void matchHeadersOnSheet(Sheet sheet, List<Map.Entry<String, ImportInformation>> entries) throws IOException {
        sheet.openStream()
                .limit(HEADER_SEARCH_DEPTH)
                .forEach(row -> row.stream()
                        .filter(Objects::nonNull)
                        .forEach(cell -> matchCell(cell, sheet.getName(), entries)));
    }

    private static void matchCell(Cell cell, String sheetName, List<Map.Entry<String, ImportInformation>> entries) {
        for (Map.Entry<String, ImportInformation> entry : entries) {
            List<FieldParam> sheetParams = entry.getValue().getParamMap().get(sheetName);
            if (sheetParams == null) continue;
            for (FieldParam fp : sheetParams) {
                if (fp.getHeaderName() != null && fp.getHeaderName().trim().equals(cell.getRawValue())) {
                    fp.setCurrentCellAddress(cell.getAddress());
                    entry.getValue().getRequiredFields().remove(FieldNameUtils.getFieldName(fp.getFieldName()));
                }
            }
        }
    }

    static void loopObjectFill(ReadableWorkbook readableWB, ExcelImportParamCore importParam) {
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
                                        fillObjectLoop(cells, paramMaps, fillObj, importParam.getWarnings());
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

    private static long getCount(Row cells) {
        return cells.stream()
                .filter(Objects::nonNull)
                .filter(cell -> cell.getType() != CellType.EMPTY)
                .count();
    }

    private static int getSkipRow(List<FieldParam> paramMaps) {
        return paramMaps
                .stream()
                .filter(fp -> fp.getCurrentCellAddress() != null)
                .map(fieldParam -> fieldParam.getCurrentCellAddress().getRow())
                .findAny()
                .orElse(0) + 1;
    }

    private static Optional<Map.Entry<String, ImportInformation>> getStringImportInformationEntry(
            ExcelImportParamCore importParam, Sheet sheet) {
        return importParam.getParamsMap()
                .entrySet()
                .stream()
                .filter(entry -> entry.getKey().startsWith(LoopOperationType.FOR.getType()))
                .filter(entry -> entry.getValue().getParamMap().containsKey(sheet.getName()))
                .findFirst();
    }

    private static <T> void fillObjectLoop(Row cells, List<FieldParam> paramMaps, T obj, List<String> warnings) {
        for (int cellNum = 0; cellNum < cells.getCellCount(); cellNum++) {
            if (cells.getCell(cellNum) != null) {
                Cell cell = cells.getCell(cellNum);
                Optional<FieldParam> any = getFieldParam(paramMaps, cell);

                if (any.isPresent()) {
                    String keys = any.get().getFieldName();
                    if (cell.getDataFormatString() != null &&
                            cell.getDataFormatString().contains("m") &&
                            cells.getCellAsDate(cellNum).isPresent()) {
                        String fieldName = FieldNameUtils.getFieldName(keys);
                        String data = cells.getCellAsDate(cellNum).get().toString();
                        FileUtils.fillObject(obj, data, fieldName, warnings);
                    } else {
                        validateAndWrite(cell, keys, obj, warnings);
                    }
                }
            }
        }
    }

    private static Optional<FieldParam> getFieldParam(List<FieldParam> paramMaps, Cell cell) {
        return paramMaps
                .stream()
                .filter(fieldParam -> fieldParam.getCurrentCellAddress() != null &&
                        fieldParam.getCurrentCellAddress().getRow() < cell.getAddress().getRow() &&
                        fieldParam.getCurrentCellAddress().getColumn() == cell.getAddress().getColumn())
                .findAny();
    }

    private static <T> void validateAndWrite(Cell cell, String key, T object, List<String> warnings) {
        CellTypeOperation.findByType(cell.getType().name()).validateAndWrite(cell, key, object, warnings);
    }
}
