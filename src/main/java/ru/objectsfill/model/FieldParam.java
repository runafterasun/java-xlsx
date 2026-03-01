package ru.objectsfill.model;

import org.dhatim.fastexcel.reader.CellAddress;

/**
 * Описание одного маркера, найденного в шаблоне Excel.
 * Хранит полное имя маркера, его адрес в шаблоне, имя листа, заголовок столбца
 * и актуальный адрес заголовка в import-файле.
 */
public class FieldParam {

    /** Полное имя маркера в формате {@code "ключ.поле"}, например {@code "test.account"}. */
    private String fieldName;

    /** Адрес ячейки маркера в шаблоне. */
    private CellAddress cellAddress;

    /** Имя листа, на котором найден маркер. */
    private String sheetName;

    /** Текст заголовка столбца из строки выше маркера (используется для loop-блоков). */
    private String headerName;

    /** Адрес ячейки заголовка в import-файле (устанавливается при поиске маркеров в данных). */
    private CellAddress currentCellAddress;

    /**
     * Возвращает полное имя маркера в формате {@code "ключ.поле"}.
     *
     * @return полное имя маркера
     */
    public String getFieldName() {
        return fieldName;
    }

    /**
     * Устанавливает полное имя маркера.
     *
     * @param fieldName полное имя маркера
     * @return текущий экземпляр для цепочки вызовов
     */
    public FieldParam setFieldName(String fieldName) {
        this.fieldName = fieldName;
        return this;
    }

    /**
     * Возвращает адрес ячейки маркера в шаблоне.
     *
     * @return адрес ячейки
     */
    public CellAddress getCellAddress() {
        return cellAddress;
    }

    /**
     * Устанавливает адрес ячейки маркера в шаблоне.
     *
     * @param cellAddress адрес ячейки
     * @return текущий экземпляр для цепочки вызовов
     */
    public FieldParam setCellAddress(CellAddress cellAddress) {
        this.cellAddress = cellAddress;
        return this;
    }

    /**
     * Возвращает имя листа, на котором найден маркер.
     *
     * @return имя листа
     */
    public String getSheetName() {
        return sheetName;
    }

    /**
     * Устанавливает имя листа.
     *
     * @param sheetName имя листа
     * @return текущий экземпляр для цепочки вызовов
     */
    public FieldParam setSheetName(String sheetName) {
        this.sheetName = sheetName;
        return this;
    }

    /**
     * Возвращает текст заголовка столбца из строки выше маркера.
     *
     * @return текст заголовка
     */
    public String getHeaderName() {
        return headerName;
    }

    /**
     * Устанавливает текст заголовка столбца.
     *
     * @param headerName текст заголовка
     * @return текущий экземпляр для цепочки вызовов
     */
    public FieldParam setHeaderName(String headerName) {
        this.headerName = headerName;
        return this;
    }

    /**
     * Возвращает адрес ячейки заголовка в import-файле.
     *
     * @return актуальный адрес заголовка
     */
    public CellAddress getCurrentCellAddress() {
        return currentCellAddress;
    }

    /**
     * Устанавливает адрес ячейки заголовка в import-файле.
     *
     * @param currentCellAddress актуальный адрес заголовка
     * @return текущий экземпляр для цепочки вызовов
     */
    public FieldParam setCurrentCellAddress(CellAddress currentCellAddress) {
        this.currentCellAddress = currentCellAddress;
        return this;
    }
}
