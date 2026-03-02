package ru.objectsfill.reader.dto;

/**
 * DTO для десериализации координат ячейки из JSON-шаблона.
 * Соответствует {@link org.dhatim.fastexcel.reader.CellAddress}.
 */
public class CellAddressDto {

    /** Номер строки (0-based). */
    private int row;

    /** Номер столбца (0-based). */
    private int col;

    public int getRow() {
        return row;
    }

    public void setRow(int row) {
        this.row = row;
    }

    public int getCol() {
        return col;
    }

    public void setCol(int col) {
        this.col = col;
    }
}
