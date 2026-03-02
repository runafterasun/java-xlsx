package ru.objectsfill;

import org.junit.jupiter.api.Test;
import ru.objectsfill.dto.ExcelImport;
import ru.objectsfill.dto.LoopDate;
import ru.objectsfill.enums.BindingMode;
import ru.objectsfill.exception.ExcelImportException;
import ru.objectsfill.model.ExcelImportParamCore;
import ru.objectsfill.model.FieldParam;
import ru.objectsfill.model.ImportInformation;
import ru.objectsfill.reader.JsonTemplateReader;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

class JsonTemplateReaderTest {

    // --- validation tests ---

    @Test
    void validation_missingFieldName_throwsWithIndex() {
        String json = """
                {"entries": [{"sheetName": "Sheet1", "cellAddress": {"row": 0, "col": 0}}]}
                """;
        ExcelImportException ex = assertThrows(ExcelImportException.class,
                () -> read(json, param()));
        assertTrue(ex.getMessage().contains("entry[0]"), ex.getMessage());
        assertTrue(ex.getMessage().contains("fieldName"), ex.getMessage());
    }

    @Test
    void validation_missingSheetName_throwsWithIndex() {
        String json = """
                {"entries": [{"fieldName": "test.account", "cellAddress": {"row": 0, "col": 0}}]}
                """;
        ExcelImportException ex = assertThrows(ExcelImportException.class,
                () -> read(json, param()));
        assertTrue(ex.getMessage().contains("entry[0]"), ex.getMessage());
        assertTrue(ex.getMessage().contains("sheetName"), ex.getMessage());
    }

    @Test
    void validation_missingBothHeaderAndCellAddress_throwsWithIndex() {
        String json = """
                {"entries": [{"fieldName": "test.account", "sheetName": "Sheet1"}]}
                """;
        ExcelImportException ex = assertThrows(ExcelImportException.class,
                () -> read(json, param()));
        assertTrue(ex.getMessage().contains("entry[0]"), ex.getMessage());
        assertTrue(ex.getMessage().contains("headerName"), ex.getMessage());
        assertTrue(ex.getMessage().contains("cellAddress"), ex.getMessage());
    }

    @Test
    void validation_secondEntryInvalid_indexIsCorrect() {
        String json = """
                {"entries": [
                  {"fieldName": "test.account", "sheetName": "Sheet1", "cellAddress": {"row": 0, "col": 0}},
                  {"fieldName": "test.currency", "sheetName": "Sheet1"}
                ]}
                """;
        ExcelImportException ex = assertThrows(ExcelImportException.class,
                () -> read(json, param()));
        assertTrue(ex.getMessage().contains("entry[1]"), ex.getMessage());
    }

    @Test
    void validation_emptyEntries_throws() {
        String json = """
                {"entries": []}
                """;
        ExcelImportException ex = assertThrows(ExcelImportException.class,
                () -> read(json, param()));
        assertTrue(ex.getMessage().contains("missing or empty"), ex.getMessage());
    }

    @Test
    void validation_invalidJson_throws() {
        ExcelImportException ex = assertThrows(ExcelImportException.class,
                () -> read("not-json", param()));
        assertTrue(ex.getMessage().contains("Failed to parse"), ex.getMessage());
    }

    // --- FieldParam construction tests ---

    @Test
    void read_headerNameOnly_fieldParamHasHeaderName() {
        String json = """
                {"entries": [
                  {"fieldName": "for.dateList.rate", "sheetName": "Sheet1", "headerName": "Ставка %"}
                ]}
                """;
        ExcelImportParamCore p = param();
        List<FieldParam> out = read(json, p);

        assertEquals(1, out.size());
        assertEquals("for.dateList.rate", out.get(0).getFieldName());
        assertEquals("Sheet1", out.get(0).getSheetName());
        assertEquals("Ставка %", out.get(0).getHeaderName());
        assertNull(out.get(0).getCellAddress());
    }

    @Test
    void read_cellAddressOnly_fieldParamHasCellAddress() {
        String json = """
                {"entries": [
                  {"fieldName": "test.account", "sheetName": "Sheet1", "cellAddress": {"row": 2, "col": 3}}
                ]}
                """;
        List<FieldParam> out = read(json, param());

        assertEquals(1, out.size());
        assertNull(out.get(0).getHeaderName());
        assertNotNull(out.get(0).getCellAddress());
        assertEquals(2, out.get(0).getCellAddress().getRow());
        assertEquals(3, out.get(0).getCellAddress().getColumn());
    }

    @Test
    void read_bothHeaderAndCellAddress_headerNameTakesPriority() {
        String json = """
                {"entries": [
                  {"fieldName": "for.dateList.rate", "sheetName": "Sheet1",
                   "headerName": "Ставка %", "cellAddress": {"row": 4, "col": 2}}
                ]}
                """;
        List<FieldParam> out = read(json, param());

        assertEquals("Ставка %", out.get(0).getHeaderName());
        assertNotNull(out.get(0).getCellAddress());
    }

    // --- BindingMode auto-detection tests ---

    @Test
    void bindingMode_headerNamePresent_setsHeaderMode() {
        String json = """
                {"entries": [
                  {"fieldName": "for.dateList.rate", "sheetName": "Sheet1", "headerName": "Ставка %"},
                  {"fieldName": "for.dateList.date", "sheetName": "Sheet1", "headerName": "Дата"}
                ]}
                """;
        ExcelImportParamCore p = param();
        read(json, p);

        assertEquals(BindingMode.HEADER, p.getParamsMap().get("for.dateList").getBindingMode());
    }

    @Test
    void bindingMode_cellAddressOnly_setsPositionMode() {
        String json = """
                {"entries": [
                  {"fieldName": "for.dateList.rate", "sheetName": "Sheet1", "cellAddress": {"row": 4, "col": 2}},
                  {"fieldName": "for.dateList.date", "sheetName": "Sheet1", "cellAddress": {"row": 4, "col": 1}}
                ]}
                """;
        ExcelImportParamCore p = param();
        read(json, p);

        assertEquals(BindingMode.POSITION, p.getParamsMap().get("for.dateList").getBindingMode());
    }

    @Test
    void bindingMode_singleObjectEntry_notAffected() {
        String json = """
                {"entries": [
                  {"fieldName": "test.account", "sheetName": "Sheet1", "cellAddress": {"row": 0, "col": 0}}
                ]}
                """;
        ExcelImportParamCore p = param();
        read(json, p);

        // Single object binding mode stays at default (HEADER is default, but mode doesn't affect singles)
        assertEquals(BindingMode.HEADER, p.getParamsMap().get("test").getBindingMode());
    }

    // --- backward compatibility test ---

    @Test
    void backwardCompat_oldSignatureStillWorks() throws Exception {
        var importParam = new ExcelImportParamCore();
        importParam.getParamsMap().put("test", new ImportInformation().setClazz(ExcelImport.class));
        importParam.getParamsMap().put("for.dateList", new ImportInformation().setClazz(LoopDate.class));

        try (InputStream tmpl = resource("02_RateModNotice_tmpl.xlsx");
             InputStream data = resource("01_RateModNotice_tmpl.xlsx")) {
            ExcelImportUtil.importExcel(importParam, tmpl, data);
        }

        assertFalse(importParam.getParamsMap().get("for.dateList").getLoopLst().isEmpty());
    }

    // --- helpers ---

    private static ExcelImportParamCore param() {
        var p = new ExcelImportParamCore();
        p.getParamsMap().put("test", new ImportInformation().setClazz(ExcelImport.class));
        p.getParamsMap().put("for.dateList", new ImportInformation().setClazz(LoopDate.class));
        return p;
    }

    private static List<FieldParam> read(String json, ExcelImportParamCore param) {
        List<FieldParam> out = new ArrayList<>();
        new JsonTemplateReader(new ByteArrayInputStream(json.getBytes(StandardCharsets.UTF_8)))
                .read(param, out);
        return out;
    }

    private InputStream resource(String filename) {
        return getClass().getResourceAsStream("/xlstemplates/" + filename);
    }
}
