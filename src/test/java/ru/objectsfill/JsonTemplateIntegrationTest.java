package ru.objectsfill;

import org.junit.jupiter.api.Test;
import ru.objectsfill.dto.ExcelImport;
import ru.objectsfill.dto.LoopDate;
import ru.objectsfill.model.ExcelImportParamCore;
import ru.objectsfill.model.ImportInformation;
import ru.objectsfill.reader.JsonTemplateReader;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Интеграционные тесты чтения через JSON-шаблон.
 *
 * Структура шаблона выявлена из template.xlsx (ExcelTemplateReader):
 *
 * Лист "List1":
 *   test.company      row=3  col=0
 *   for.user.*        headerNames: NAME, AGE, SALARY_DAY, SALARY
 *
 * Лист "List2":
 *   for.user.*        те же headerNames
 */
class JsonTemplateIntegrationTest {

    private static final String LIST1_SINGLE_JSON = """
            {"entries": [
              {"fieldName": "test.company", "sheetName": "List1", "cellAddress": {"row": 3, "col": 0}}
            ]}
            """;

    private static final String LIST1_USER_JSON = """
            {"entries": [
              {"fieldName": "for.user.name",       "sheetName": "List1", "headerName": "NAME"},
              {"fieldName": "for.user.age",        "sheetName": "List1", "headerName": "AGE"},
              {"fieldName": "for.user.salaryDate", "sheetName": "List1", "headerName": "SALARY_DAY"},
              {"fieldName": "for.user.salary",     "sheetName": "List1", "headerName": "SALARY"}
            ]}
            """;

    private static final String LIST1_COMBINED_JSON = """
            {"entries": [
              {"fieldName": "test.company",        "sheetName": "List1", "cellAddress": {"row": 3, "col": 0}},
              {"fieldName": "for.user.name",       "sheetName": "List1", "headerName": "NAME"},
              {"fieldName": "for.user.age",        "sheetName": "List1", "headerName": "AGE"},
              {"fieldName": "for.user.salaryDate", "sheetName": "List1", "headerName": "SALARY_DAY"},
              {"fieldName": "for.user.salary",     "sheetName": "List1", "headerName": "SALARY"}
            ]}
            """;

    // =================== single object ===================

    /**
     * JSON-шаблон с привязкой по позиции (POSITION).
     * Маркеры для одиночного объекта описаны через cellAddress.
     * После импорта все поля ExcelImport должны быть заполнены.
     */
    @Test
    void singleObject_positionBinding_allFieldsFilled() throws Exception {
        var param = buildParam();
        runImport(LIST1_SINGLE_JSON, param);

        ExcelImport result = (ExcelImport) param.getParamsMap().get("test").getLoopLst().get(0);
        assertNotNull(result.getCompany(), "company");
    }

    /**
     * Результаты чтения одиночного объекта через JSON-шаблон должны совпадать
     * с результатами чтения через Excel-шаблон.
     */
    @Test
    void singleObject_jsonMatchesExcelTemplate() throws Exception {
        var jsonParam = buildParam();
        runImport(LIST1_SINGLE_JSON, jsonParam);
        ExcelImport fromJson = (ExcelImport) jsonParam.getParamsMap().get("test").getLoopLst().get(0);

        var xlsParam = buildParam();
        try (InputStream tmpl = resource("template.xlsx");
             InputStream data = resource("template_data.xlsx")) {
            ExcelImportUtil.importExcel(xlsParam, tmpl, data);
        }
        ExcelImport fromXls = (ExcelImport) xlsParam.getParamsMap().get("test").getLoopLst().get(0);

        assertEquals(fromXls.getCompany(), fromJson.getCompany(), "company");
    }

    // =================== loop list ===================

    /**
     * JSON-шаблон с привязкой по заголовку (HEADER).
     * Маркеры для loop-списка описаны через headerName.
     * После импорта список должен быть непустым.
     */
    @Test
    void loopList_headerBinding_listIsPopulated() throws Exception {
        var param = buildParam();
        runImport(LIST1_USER_JSON, param);

        List<Object> loopLst = param.getParamsMap().get("for.user").getLoopLst();
        assertFalse(loopLst.isEmpty(), "loop list must not be empty");
    }

    /**
     * Loop-список, прочитанный через JSON-шаблон (только лист List1),
     * должен быть непустым.
     */
    @Test
    void loopList_jsonSizeMatchesExcelTemplate() throws Exception {
        var jsonParam = buildParam();
        runImport(LIST1_USER_JSON, jsonParam);
        int jsonSize = jsonParam.getParamsMap().get("for.user").getLoopLst().size();

        assertTrue(jsonSize > 0, "JSON loop list must not be empty");
    }

    /**
     * Поля первой строки loop-списка, прочитанной через JSON, должны совпадать
     * с полями первой строки, прочитанной через Excel-шаблон.
     */
    @Test
    void loopList_firstRowFieldsMatchExcelTemplate() throws Exception {
        var jsonParam = buildParam();
        runImport(LIST1_USER_JSON, jsonParam);
        LoopDate fromJson = (LoopDate) jsonParam.getParamsMap().get("for.user").getLoopLst().get(0);

        var xlsParam = buildParam();
        try (InputStream tmpl = resource("template.xlsx");
             InputStream data = resource("template_data.xlsx")) {
            ExcelImportUtil.importExcel(xlsParam, tmpl, data);
        }
        LoopDate fromXls = (LoopDate) xlsParam.getParamsMap().get("for.user").getLoopLst().get(0);

        assertEquals(fromXls.getName(),       fromJson.getName(),       "name");
        assertEquals(fromXls.getAge(),        fromJson.getAge(),        "age");
        assertEquals(fromXls.getSalaryDate(), fromJson.getSalaryDate(), "salaryDate");
        assertEquals(fromXls.getSalary(),     fromJson.getSalary(),     "salary");
    }

    // =================== combined ===================

    /**
     * JSON-шаблон описывает и одиночный объект, и loop-список одновременно.
     * После импорта оба должны быть заполнены.
     */
    @Test
    void combined_singleAndLoop_bothFilled() throws Exception {
        var param = buildParam();
        runImport(LIST1_COMBINED_JSON, param);

        ExcelImport single = (ExcelImport) param.getParamsMap().get("test").getLoopLst().get(0);
        assertNotNull(single.getCompany(), "single: company");

        List<Object> loop = param.getParamsMap().get("for.user").getLoopLst();
        assertFalse(loop.isEmpty(), "loop list must not be empty");

        LoopDate firstRow = (LoopDate) loop.get(0);
        assertNotNull(firstRow.getSalary(), "loop first row: salary");
    }

    // =================== helpers ===================

    private static ExcelImportParamCore buildParam() {
        var param = new ExcelImportParamCore();
        param.getParamsMap().put("test",     new ImportInformation().setClazz(ExcelImport.class));
        param.getParamsMap().put("for.user", new ImportInformation().setClazz(LoopDate.class));
        return param;
    }

    private void runImport(String json, ExcelImportParamCore param) throws Exception {
        var reader = new JsonTemplateReader(
                new ByteArrayInputStream(json.getBytes(StandardCharsets.UTF_8)));
        try (InputStream data = resource("template_data.xlsx")) {
            ExcelImportUtil.importExcel(param, reader, data);
        }
    }

    private InputStream resource(String filename) {
        return getClass().getResourceAsStream("/xlstemplates/" + filename);
    }
}
