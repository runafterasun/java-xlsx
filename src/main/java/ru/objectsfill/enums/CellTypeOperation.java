package ru.objectsfill.enums;

import org.dhatim.fastexcel.reader.Cell;
import ru.objectsfill.exception.ExcelImportException;
import ru.objectsfill.util.FieldNameUtils;

import java.math.BigDecimal;
import java.util.Arrays;
import java.util.List;

import static ru.objectsfill.util.FileUtils.fillObject;


public enum CellTypeOperation {

    NUMBER("NUMBER") {
        @Override
        public <T> void validateAndWrite(Cell cell, String key, T object, List<String> warnings) {
            try {
                BigDecimal cellValue = new BigDecimal(cell.getRawValue());
                writeCell(cell, cellValue.toPlainString(), key, object, warnings);
            } catch (NumberFormatException e) {
                throw new ExcelImportException(
                        "Failed to parse numeric value '" + cell.getRawValue() + "' at " + cell.getAddress(), e);
            }
        }
    },
    DEFAULT("DEFAULT") {
        @Override
        public <T> void validateAndWrite(Cell cell, String key, T object, List<String> warnings) {
            writeCell(cell, cell.getRawValue(), key, object, warnings);
        }
    };

    CellTypeOperation(String type) {
        this.type = type;
    }

    private final String type;

    public String getType() {
        return type;
    }

    public abstract <T> void validateAndWrite(Cell cell, String key, T object, List<String> warnings);

    /**
     * Ищет константу перечисления по строковому типу ячейки (без учёта регистра).
     * Если тип не найден, возвращает {@link #DEFAULT}.
     *
     * @param cellType строковое представление типа ячейки (например, {@code "NUMBER"})
     * @return соответствующая константа, или {@link #DEFAULT}, если совпадений нет
     */
    public static CellTypeOperation findByType(String cellType) {
        return Arrays.stream(CellTypeOperation.values())
                .filter(cellTypeOperation -> cellTypeOperation.getType().equals(cellType.toUpperCase()))
                .findFirst()
                .orElse(DEFAULT);
    }

    private static <T> void writeCell(Cell cell, String value, String key, T object, List<String> warnings) {
        fillObject(object, value, FieldNameUtils.getFieldName(key), warnings);
        fillObject(object, cell.getAddress().getRow() + 1, LoopOperationType.ROW.getType(), warnings);
    }
}
