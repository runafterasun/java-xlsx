package ru.objectsfill.enums;

import org.dhatim.fastexcel.reader.Cell;

import java.math.BigDecimal;
import java.util.Arrays;

import static ru.objectsfill.util.FileUtils.fillObject;


public enum CellTypeOperation {

    NUMBER("NUMBER") {
        @Override
        public <T> void validateAndWrite(Cell cell, String key, T object) {
            var cellValue = new BigDecimal(cell.getRawValue());
            var fieldName = getSubstringKey(key);
            fillObject(object, cellValue.toPlainString(), fieldName);
            fillObject(object, cell.getAddress().getRow() + 1, LoopOperationType.ROW.getType());
        }
    },
    DEFAULT("DEFAULT") {
        @Override
        public <T> void validateAndWrite(Cell cell, String key, T object) {
            var fieldName = getSubstringKey(key);
            fillObject(object, cell.getRawValue(), fieldName);
            fillObject(object, cell.getAddress().getRow() + 1, LoopOperationType.ROW.getType());
        }
    };

    CellTypeOperation(String type) {
        this.type = type;
    }

    private final String type;

    public String getType() {
        return type;
    }

    public abstract <T> void validateAndWrite(Cell cell, String key, T object);

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

    /**
     * Извлекает имя поля из полного имени маркера (часть после последней точки).
     * Например, из {@code "test.account"} вернёт {@code "account"}.
     *
     * @param key полное имя маркера
     * @return имя поля
     */
    public static String getSubstringKey(String key) {
        return key.substring(key.lastIndexOf(".") + 1);
    }
}
