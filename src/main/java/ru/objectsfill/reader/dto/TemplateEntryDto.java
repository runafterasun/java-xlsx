package ru.objectsfill.reader.dto;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;

/**
 * DTO для одной записи JSON-шаблона.
 * Описывает один маркер: имя поля, лист, заголовок или позицию.
 *
 * <p>Правила заполнения:
 * <ul>
 *   <li>{@code fieldName} — обязательно.</li>
 *   <li>{@code sheetName} — обязательно.</li>
 *   <li>Хотя бы одно из {@code headerName} или {@code cellAddress} — обязательно.</li>
 *   <li>Если {@code headerName} присутствует — режим привязки {@code HEADER}.</li>
 *   <li>Если только {@code cellAddress} — режим привязки {@code POSITION}.</li>
 * </ul>
 */
@JsonIgnoreProperties(ignoreUnknown = true)
public class TemplateEntryDto {

    /** Полное имя маркера, например {@code "test.account"} или {@code "for.dateList.rate"}. */
    private String fieldName;

    /** Имя листа, на котором расположен маркер. */
    private String sheetName;

    /**
     * Текст заголовка столбца (строка над строкой данных).
     * При наличии — используется режим {@code HEADER}.
     */
    private String headerName;

    /**
     * Координаты ячейки маркера (0-based row/col).
     * Используется как резервный вариант, если {@code headerName} не задан.
     */
    private CellAddressDto cellAddress;

    public String getFieldName() {
        return fieldName;
    }

    public void setFieldName(String fieldName) {
        this.fieldName = fieldName;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public String getHeaderName() {
        return headerName;
    }

    public void setHeaderName(String headerName) {
        this.headerName = headerName;
    }

    public CellAddressDto getCellAddress() {
        return cellAddress;
    }

    public void setCellAddress(CellAddressDto cellAddress) {
        this.cellAddress = cellAddress;
    }
}
