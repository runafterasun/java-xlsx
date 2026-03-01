package ru.objectsfill;

import lombok.AllArgsConstructor;
import lombok.Getter;
import org.dhatim.fastexcel.reader.Cell;

import java.math.BigDecimal;
import java.util.Arrays;

import static ru.objectsfill.FileUtils.fillObject;


@Getter
@AllArgsConstructor
public enum CellTypeOperation {

    NUMBER("NUMBER") {
        @Override
        public <T> void validateAndWrite(Cell cell, String key, T object) {
            var cellValue = new BigDecimal(cell.getRawValue());
            var fieldName = getSubstringKey(key);
            fillObject(object, cellValue.toPlainString(), fieldName);
            fillObject(object, cell.getAddress().getRow(), LoopOperationType.ROW.getType());
        }
    },
    DEFAULT("DEFAULT") {
        @Override
        public <T> void validateAndWrite(Cell cell, String key, T object) {
            var fieldName = getSubstringKey(key);
            fillObject(object, cell.getRawValue(), fieldName);
            fillObject(object, cell.getAddress().getRow(), LoopOperationType.ROW.getType());
        }
    };

    private final String type;

    public abstract <T> void validateAndWrite(Cell cell, String key, T object);

    public static CellTypeOperation findByType(String cellType) {
        return Arrays.stream(CellTypeOperation.values())
                .filter(cellTypeOperation -> cellTypeOperation.getType().equals(cellType.toUpperCase()))
                .findFirst()
                .orElse(DEFAULT);
    }

    private static String getSubstringKey(String key) {
        return key.substring(key.lastIndexOf(".") + 1);
    }


}
