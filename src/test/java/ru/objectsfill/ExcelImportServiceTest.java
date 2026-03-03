package ru.objectsfill;

import org.junit.jupiter.api.Test;
import ru.objectsfill.dto.ExcelImport;
import ru.objectsfill.dto.LoopDate;
import ru.objectsfill.enums.BindingMode;
import ru.objectsfill.model.ExcelImportParamCore;
import ru.objectsfill.exception.ExcelImportException;
import ru.objectsfill.model.FieldParam;
import ru.objectsfill.model.ImportInformation;
import ru.objectsfill.reader.JsonTemplateReader;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.math.BigDecimal;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

class ExcelImportServiceTest {

    @Test
    void sheet0_singleObjectIsFilled() throws Exception {
        var importParam = new ExcelImportParamCore();

        importParam.getParamsMap().put("test",     new ImportInformation().setClazz(ExcelImport.class));
        importParam.getParamsMap().put("for.user", new ImportInformation().setClazz(LoopDate.class));

        runImport(importParam);

        ExcelImport excelImport = (ExcelImport) importParam.getParamsMap().get("test").getLoopLst().get(0);
        assertNotNull(excelImport.getCompany(), "company");
    }

    @Test
    void sheet0_loopListIsPopulated() throws Exception {
        var importParam = new ExcelImportParamCore();

        importParam.getParamsMap().put("test",     new ImportInformation().setClazz(ExcelImport.class));
        importParam.getParamsMap().put("for.user", new ImportInformation().setClazz(LoopDate.class));

        runImport(importParam);

        assertFalse(importParam.getParamsMap().get("for.user").getLoopLst().isEmpty(), "loop list should not be empty");
    }

    // --- BindingMode tests ---

    @Test
    void bindingMode_defaultIsHeader() {
        var info = new ImportInformation();
        assertEquals(BindingMode.HEADER, info.getBindingMode());
    }

    @Test
    void bindingMode_explicitHeader_behavesLikeDefault() throws Exception {
        var importParam = new ExcelImportParamCore();

        importParam.getParamsMap().put("test", new ImportInformation().setClazz(ExcelImport.class));
        importParam.getParamsMap().put("for.user",
                new ImportInformation().setClazz(LoopDate.class).setBindingMode(BindingMode.HEADER));

        runImport(importParam);

        assertFalse(importParam.getParamsMap().get("for.user").getLoopLst().isEmpty(),
                "HEADER mode: loop list should not be empty");
    }

    @Test
    void bindingMode_position_currentCellAddressEqualsTemplateAddress() throws Exception {
        var importParam = new ExcelImportParamCore();

        importParam.getParamsMap().put("test", new ImportInformation().setClazz(ExcelImport.class));
        importParam.getParamsMap().put("for.user",
                new ImportInformation().setClazz(LoopDate.class).setBindingMode(BindingMode.POSITION));

        try (InputStream tmpl = resource("template.xlsx");
             InputStream data = resource("template_data.xlsx")) {
            ExcelImportUtil.importExcel(importParam, tmpl, data);
        }

        List<FieldParam> fieldParams = new ArrayList<>(
                importParam.getParamsMap().get("for.user").getParamMap().values()
                        .stream().findFirst().orElseThrow());

        for (FieldParam fp : fieldParams) {
            assertEquals(fp.getCellAddress(), fp.getCurrentCellAddress(),
                    "POSITION mode: currentCellAddress must equal cellAddress for marker: " + fp.getFieldName());
        }
    }

    @Test
    void bindingMode_position_requiredFieldsAreCleared() throws Exception {
        var importParam = new ExcelImportParamCore();

        importParam.getParamsMap().put("test", new ImportInformation().setClazz(ExcelImport.class));

        var loopInfo = new ImportInformation()
                .setClazz(LoopDate.class)
                .setBindingMode(BindingMode.POSITION)
                .setRequiredFields(new ArrayList<>(List.of("salary", "age")));
        importParam.getParamsMap().put("for.user", loopInfo);

        try (InputStream tmpl = resource("template.xlsx");
             InputStream data = resource("template_data.xlsx")) {
            ExcelImportUtil.importExcel(importParam, tmpl, data);
        }

        assertTrue(loopInfo.getRequiredFields().isEmpty(),
                "POSITION mode: requiredFields must be cleared after binding");
    }

    // --- Warnings tests ---

    @Test
    void warnings_unknownField_warningAdded() throws Exception {
        String json = """
                {"entries": [
                  {"fieldName": "test.nonExistentField", "sheetName": "List1", "cellAddress": {"row": 3, "col": 0}}
                ]}
                """;
        var param = new ExcelImportParamCore();
        param.getParamsMap().put("test", new ImportInformation().setClazz(ExcelImport.class));

        var reader = new JsonTemplateReader(new ByteArrayInputStream(json.getBytes(StandardCharsets.UTF_8)));
        try (InputStream data = resource("template_data.xlsx")) {
            ExcelImportUtil.importExcel(param, reader, data);
        }

        assertFalse(param.getWarnings().isEmpty(), "warnings should contain field-not-found entries");
        assertTrue(param.getWarnings().stream().anyMatch(w -> w.contains("nonExistentField")),
                "warning should mention the missing field name");
    }

    @Test
    void warnings_unknownHeader_loopListEmptyNoException() throws Exception {
        String json = """
                {"entries": [
                  {"fieldName": "for.user.name", "sheetName": "List1", "headerName": "UNKNOWN_COLUMN_HEADER"}
                ]}
                """;
        var param = new ExcelImportParamCore();
        param.getParamsMap().put("for.user", new ImportInformation().setClazz(LoopDate.class));

        var reader = new JsonTemplateReader(new ByteArrayInputStream(json.getBytes(StandardCharsets.UTF_8)));
        try (InputStream data = resource("template_data.xlsx")) {
            assertDoesNotThrow(() -> ExcelImportUtil.importExcel(param, reader, data));
        }

        assertTrue(param.getParamsMap().get("for.user").getLoopLst().isEmpty(),
                "loop list must be empty when header is not found");
    }

    // --- validateHeaderBindingHasHeaderName ---

    @Test
    void headerBinding_missingHeaderName_throwsException() {
        String json = """
                {"entries": [
                  {"fieldName": "for.user.salary", "sheetName": "List1", "headerName": "SALARY"},
                  {"fieldName": "for.user.age",    "sheetName": "List1", "cellAddress": {"row": 4, "col": 5}}
                ]}
                """;
        var param = new ExcelImportParamCore();
        param.getParamsMap().put("for.user", new ImportInformation().setClazz(LoopDate.class));

        var reader = new JsonTemplateReader(new ByteArrayInputStream(json.getBytes(StandardCharsets.UTF_8)));

        ExcelImportException ex = assertThrows(ExcelImportException.class,
                () -> ExcelImportUtil.importExcel(param, reader, InputStream.nullInputStream()));
        assertTrue(ex.getMessage().contains("HEADER binding"), ex.getMessage());
    }

    // --- числовые ячейки ---

    @Test
    void numberCell_salaryStoredAsPlainDecimal() throws Exception {
        var importParam = new ExcelImportParamCore();
        importParam.getParamsMap().put("for.user", new ImportInformation().setClazz(LoopDate.class));
        runImport(importParam);

        List<Object> list = importParam.getParamsMap().get("for.user").getLoopLst();
        assertFalse(list.isEmpty(), "нужны строки данных для проверки числового парсинга");

        list.stream()
                .map(o -> (LoopDate) o)
                .map(LoopDate::getSalary)
                .filter(s -> s != null && !s.isBlank())
                .forEach(salary -> {
                    assertDoesNotThrow(() -> new BigDecimal(salary),
                            "значение salary должно быть корректным десятичным числом: " + salary);
                    assertFalse(salary.toUpperCase().contains("E"),
                            "salary не должен быть в научной нотации (toPlainString): " + salary);
                });
    }

    // --- несколько листов ---

    @Test
    void multiSheet_markersFoundOnBothSheets() throws Exception {
        var importParam = new ExcelImportParamCore();
        importParam.getParamsMap().put("for.user", new ImportInformation().setClazz(LoopDate.class));
        runImport(importParam);

        Map<String, List<FieldParam>> paramMap =
                importParam.getParamsMap().get("for.user").getParamMap();
        assertTrue(paramMap.containsKey("List1"), "маркеры должны быть найдены на листе List1");
        assertTrue(paramMap.containsKey("List2"), "маркеры должны быть найдены на листе List2");
    }

    @Test
    void multiSheet_excelTemplateGivesMoreRowsThanSingleSheet() throws Exception {
        var xlsParam = new ExcelImportParamCore();
        xlsParam.getParamsMap().put("for.user", new ImportInformation().setClazz(LoopDate.class));
        runImport(xlsParam);
        int excelRows = xlsParam.getParamsMap().get("for.user").getLoopLst().size();

        String singleSheetJson = """
                {"entries": [
                  {"fieldName": "for.user.name",       "sheetName": "List1", "headerName": "NAME"},
                  {"fieldName": "for.user.age",        "sheetName": "List1", "headerName": "AGE"},
                  {"fieldName": "for.user.salaryDate", "sheetName": "List1", "headerName": "SALARY_DAY"},
                  {"fieldName": "for.user.salary",     "sheetName": "List1", "headerName": "SALARY"}
                ]}
                """;
        var jsonParam = new ExcelImportParamCore();
        jsonParam.getParamsMap().put("for.user", new ImportInformation().setClazz(LoopDate.class));
        var reader = new JsonTemplateReader(new ByteArrayInputStream(singleSheetJson.getBytes(StandardCharsets.UTF_8)));
        try (InputStream data = resource("template_data.xlsx")) {
            ExcelImportUtil.importExcel(jsonParam, reader, data);
        }
        int jsonRows = jsonParam.getParamsMap().get("for.user").getLoopLst().size();

        assertTrue(excelRows > jsonRows,
                "Excel-шаблон (два листа) должен дать больше строк, чем JSON (один лист): "
                        + excelRows + " vs " + jsonRows);
    }

    private void runImport(ExcelImportParamCore param) throws Exception {
        try (InputStream tmpl = resource("template.xlsx");
             InputStream data = resource("template_data.xlsx")) {
            ExcelImportUtil.importExcel(param, tmpl, data);
        }
    }

    private InputStream resource(String filename) {
        return getClass().getResourceAsStream("/xlstemplates/" + filename);
    }
}
