package ru.objectsfill;

import org.dhatim.fastexcel.reader.*;
import ru.objectsfill.enums.CellTypeOperation;
import ru.objectsfill.enums.LoopOperationType;
import ru.objectsfill.exception.ExcelImportException;
import ru.objectsfill.model.ExcelImportParamCore;
import ru.objectsfill.model.FieldParam;

import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.function.Function;
import java.util.stream.Collectors;

class SingleObjectFiller {

    private SingleObjectFiller() {}

    private static final int HEADER_SEARCH_DEPTH = 100;

    static void singleObjectFill(ReadableWorkbook readableWB, ExcelImportParamCore importParam) {
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
                                            .limit(HEADER_SEARCH_DEPTH)
                                            .forEach(ss -> ss.stream()
                                                    .filter(Objects::nonNull)
                                                    .forEach(cell -> {
                                                        if (collect.containsKey(cell.getAddress())) {
                                                            FieldParam fieldParam = collect.get(cell.getAddress());
                                                            validateAndWrite(cell, fieldParam.getFieldName(), object, importParam.getWarnings());
                                                        }
                                                    }));
                                    valMap.getValue().getLoopLst().add(object);
                                } catch (ExcelImportException e) {
                                    throw e;
                                } catch (Exception e) {
                                    throw new ExcelImportException(
                                            "Failed to instantiate or fill object for key: " + valMap.getKey(), e);
                                }
                            }
                        }));
    }

    private static <T> void validateAndWrite(Cell cell, String key, T object, List<String> warnings) {
        CellTypeOperation.findByType(cell.getType().name()).validateAndWrite(cell, key, object, warnings);
    }
}
