package ru.objectsfill;

import org.junit.jupiter.api.Test;
import ru.objectsfill.dto.ExcelImport;
import ru.objectsfill.dto.LoopDate;
import ru.objectsfill.model.ExcelExportParamCore;
import ru.objectsfill.model.ExcelImportParamCore;
import ru.objectsfill.model.ExportInformation;
import ru.objectsfill.model.ImportInformation;
import ru.objectsfill.reader.JsonTemplateReader;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

class ExcelExportUtilTest {

    // JSON-шаблон: одиночный объект по позиции
    private static final String SINGLE_JSON = """
            {"entries": [
              {"fieldName": "test.company", "sheetName": "Sheet1", "cellAddress": {"row": 0, "col": 0}}
            ]}
            """;

    // JSON-шаблон: loop с заголовками.
    // cellAddress.row=1 — строка данных (вариант A: маркерная строка заменяется данными).
    // row=0 — строка заголовков (headerName).
    // При экспорте: заголовки пишутся в row=0, данные в row=1,2,3,...
    // При импорте: HEADER mode — заголовок найден в row=0, данные читаются с row=1+.
    private static final String LOOP_JSON = """
            {"entries": [
              {"fieldName": "for.user.name",   "sheetName": "Sheet1", "cellAddress": {"row": 1, "col": 0}, "headerName": "NAME"},
              {"fieldName": "for.user.age",    "sheetName": "Sheet1", "cellAddress": {"row": 1, "col": 1}, "headerName": "AGE"},
              {"fieldName": "for.user.salary", "sheetName": "Sheet1", "cellAddress": {"row": 1, "col": 2}, "headerName": "SALARY"}
            ]}
            """;

    @Test
    void singleObject_exportAndReimport_valuesMatch() throws Exception {
        ExcelImport source = buildSingleObject();

        byte[] xlsx = export(SINGLE_JSON, "test", source);

        ExcelImportParamCore importParam = new ExcelImportParamCore();
        importParam.getParamsMap().put("test", new ImportInformation().setClazz(ExcelImport.class));
        ExcelImportUtil.importExcel(importParam,
                new JsonTemplateReader(jsonStream(SINGLE_JSON)),
                new ByteArrayInputStream(xlsx));

        ExcelImport result = (ExcelImport) importParam.getParamsMap().get("test").getLoopLst().get(0);
        assertEquals(source.getCompany(), result.getCompany(), "company");
    }

    @Test
    void loopObjects_exportAndReimport_allRowsPresent() throws Exception {
        List<LoopDate> sourceList = List.of(
                buildLoopDate("Alice", "30", "550"),
                buildLoopDate("Bob",   "25", "430"),
                buildLoopDate("Carol", "35", "670")
        );

        byte[] xlsx = exportLoop(LOOP_JSON, sourceList);

        ExcelImportParamCore importParam = new ExcelImportParamCore();
        importParam.getParamsMap().put("for.user", new ImportInformation().setClazz(LoopDate.class));
        ExcelImportUtil.importExcel(importParam,
                new JsonTemplateReader(jsonStream(LOOP_JSON)),
                new ByteArrayInputStream(xlsx));

        List<Object> result = importParam.getParamsMap().get("for.user").getLoopLst();
        assertEquals(sourceList.size(), result.size(), "row count");

        for (int i = 0; i < sourceList.size(); i++) {
            LoopDate src = sourceList.get(i);
            LoopDate res = (LoopDate) result.get(i);
            assertEquals(src.getName(),   res.getName(),   "name["   + i + "]");
            assertEquals(src.getAge(),    res.getAge(),    "age["    + i + "]");
            assertEquals(src.getSalary(), res.getSalary(), "salary[" + i + "]");
        }
    }

    @Test
    void unknownField_warningAdded_noException() {
        ExcelImport source = buildSingleObject();

        String jsonWithUnknown = """
                {"entries": [
                  {"fieldName": "test.nonExistent", "sheetName": "Sheet1", "cellAddress": {"row": 0, "col": 0}}
                ]}
                """;

        ExcelExportParamCore exportParam = new ExcelExportParamCore();
        exportParam.getParamsMap().put("test",
                new ExportInformation().setDataList(List.of(source)));

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        assertDoesNotThrow(() ->
                ExcelExportUtil.exportExcel(exportParam,
                        new JsonTemplateReader(jsonStream(jsonWithUnknown)), out));

        assertFalse(exportParam.getWarnings().isEmpty(), "warning should be added for unknown field");
        assertTrue(exportParam.getWarnings().stream().anyMatch(w -> w.contains("nonExistent")),
                "warning should mention the missing field name");
    }

    // --- Excel template tests ---

    @Test
    void excelTemplate_singleObject_roundtrip() throws Exception {
        // import original data
        ExcelImportParamCore origParam = buildImportParam();
        try (InputStream tmpl = resource("template.xlsx");
             InputStream data = resource("template_data.xlsx")) {
            ExcelImportUtil.importExcel(origParam, tmpl, data);
        }
        ExcelImport orig = (ExcelImport) origParam.getParamsMap().get("test").getLoopLst().get(0);

        // export
        ExcelExportParamCore exportParam = new ExcelExportParamCore();
        exportParam.getParamsMap().put("test", new ExportInformation().setDataList(List.of(orig)));
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (InputStream tmpl = resource("template.xlsx")) {
            ExcelExportUtil.exportExcel(exportParam, tmpl, out);
        }

        // re-import
        ExcelImportParamCore reimportParam = buildImportParam();
        try (InputStream tmpl = resource("template.xlsx")) {
            ExcelImportUtil.importExcel(reimportParam, tmpl, new ByteArrayInputStream(out.toByteArray()));
        }
        ExcelImport result = (ExcelImport) reimportParam.getParamsMap().get("test").getLoopLst().get(0);

        assertEquals(orig.getCompany(), result.getCompany(), "company");
    }

    @Test
    void excelTemplate_withStyles_singleObject_roundtrip() throws Exception {
        // import original data
        ExcelImportParamCore origParam = buildImportParam();
        try (InputStream tmpl = resource("template.xlsx");
             InputStream data = resource("template_data.xlsx")) {
            ExcelImportUtil.importExcel(origParam, tmpl, data);
        }
        ExcelImport orig = (ExcelImport) origParam.getParamsMap().get("test").getLoopLst().get(0);

        // export WITH styles
        ExcelExportParamCore exportParam = new ExcelExportParamCore();
        exportParam.getParamsMap().put("test", new ExportInformation().setDataList(List.of(orig)));
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (InputStream tmpl = resource("template.xlsx")) {
            ExcelExportUtil.exportExcelWithStyles(exportParam, tmpl, out);
        }
        assertTrue(out.size() > 0, "output must not be empty");
        assertTrue(exportParam.getWarnings().isEmpty(),
                "no warnings expected, got: " + exportParam.getWarnings());

        // re-import — values must survive the styled export
        ExcelImportParamCore reimportParam = buildImportParam();
        try (InputStream tmpl = resource("template.xlsx")) {
            ExcelImportUtil.importExcel(reimportParam, tmpl, new ByteArrayInputStream(out.toByteArray()));
        }
        ExcelImport result = (ExcelImport) reimportParam.getParamsMap().get("test").getLoopLst().get(0);
        assertEquals(orig.getCompany(), result.getCompany(), "company");
    }

    @Test
    void excelTemplate_withStyles_customDepth_doesNotThrow() {
        ExcelExportParamCore exportParam = new ExcelExportParamCore();
        exportParam.getParamsMap().put("test",
                new ExportInformation().setDataList(List.of(buildSingleObject())));

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        assertDoesNotThrow(() -> {
            try (InputStream tmpl = resource("template.xlsx")) {
                ExcelExportUtil.exportExcelWithStyles(exportParam, tmpl, out, 10);
            }
        });
        assertTrue(out.size() > 0);
    }

    @Test
    void excelTemplate_loopObjects_firstRowRoundtrip() throws Exception {
        // import original data
        ExcelImportParamCore origParam = buildImportParam();
        try (InputStream tmpl = resource("template.xlsx");
             InputStream data = resource("template_data.xlsx")) {
            ExcelImportUtil.importExcel(origParam, tmpl, data);
        }
        List<Object> origLoop = origParam.getParamsMap().get("for.user").getLoopLst();
        assertFalse(origLoop.isEmpty(), "original loop must not be empty");

        // export
        ExcelExportParamCore exportParam = new ExcelExportParamCore();
        exportParam.getParamsMap().put("for.user", new ExportInformation().setDataList(origLoop));
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (InputStream tmpl = resource("template.xlsx")) {
            ExcelExportUtil.exportExcel(exportParam, tmpl, out);
        }

        // re-import (template has 2 sheets with loop markers → more rows than original)
        ExcelImportParamCore reimportParam = buildImportParam();
        try (InputStream tmpl = resource("template.xlsx")) {
            ExcelImportUtil.importExcel(reimportParam, tmpl, new ByteArrayInputStream(out.toByteArray()));
        }
        List<Object> result = reimportParam.getParamsMap().get("for.user").getLoopLst();
        assertFalse(result.isEmpty(), "re-imported loop must not be empty");

        // first row must match (comes from first sheet "List1")
        LoopDate origFirst  = (LoopDate) origLoop.get(0);
        LoopDate resultFirst = (LoopDate) result.get(0);
        assertEquals(origFirst.getName(),   resultFirst.getName(),   "name[0]");
        assertEquals(origFirst.getAge(),    resultFirst.getAge(),    "age[0]");
        assertEquals(origFirst.getSalary(), resultFirst.getSalary(), "salary[0]");
    }

    // --- helpers ---

    private byte[] export(String json, String key, Object dataObject) {
        ExcelExportParamCore exportParam = new ExcelExportParamCore();
        exportParam.getParamsMap().put(key,
                new ExportInformation().setDataList(List.of(dataObject)));

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelExportUtil.exportExcel(exportParam, new JsonTemplateReader(jsonStream(json)), out);
        return out.toByteArray();
    }

    private byte[] exportLoop(String json, List<LoopDate> data) {
        ExcelExportParamCore exportParam = new ExcelExportParamCore();
        exportParam.getParamsMap().put("for.user",
                new ExportInformation().setDataList(List.copyOf(data)));

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelExportUtil.exportExcel(exportParam, new JsonTemplateReader(jsonStream(json)), out);
        return out.toByteArray();
    }

    private static ByteArrayInputStream jsonStream(String json) {
        return new ByteArrayInputStream(json.getBytes(StandardCharsets.UTF_8));
    }

    private static ExcelImportParamCore buildImportParam() {
        ExcelImportParamCore param = new ExcelImportParamCore();
        param.getParamsMap().put("test",     new ImportInformation().setClazz(ExcelImport.class));
        param.getParamsMap().put("for.user", new ImportInformation().setClazz(LoopDate.class));
        return param;
    }

    private InputStream resource(String filename) {
        return getClass().getResourceAsStream("/xlstemplates/" + filename);
    }

    private static ExcelImport buildSingleObject() {
        ExcelImport obj = new ExcelImport();
        obj.setCompany("ACME Corp");
        return obj;
    }

    private static LoopDate buildLoopDate(String name, String age, String salary) {
        LoopDate ld = new LoopDate();
        ld.setName(name);
        ld.setAge(age);
        ld.setSalary(salary);
        return ld;
    }
}
