package ru.objectsfill.reader.dto;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;

import java.util.List;

/**
 * Корневой DTO JSON-шаблона.
 *
 * <p>Пример JSON:
 * <pre>{@code
 * {
 *   "entries": [
 *     { "fieldName": "test.account",      "sheetName": "Sheet1", "cellAddress": {"row": 0, "col": 0} },
 *     { "fieldName": "for.dateList.rate", "sheetName": "Sheet1", "headerName": "Ставка %" }
 *   ]
 * }
 * }</pre>
 */
@JsonIgnoreProperties(ignoreUnknown = true)
public class TemplateDto {

    /** Список маркеров шаблона. */
    private List<TemplateEntryDto> entries;

    public List<TemplateEntryDto> getEntries() {
        return entries;
    }

    public void setEntries(List<TemplateEntryDto> entries) {
        this.entries = entries;
    }
}
