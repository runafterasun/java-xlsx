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
              {"fieldName": "test.account",       "sheetName": "Sheet1", "cellAddress": {"row": 0, "col": 0}},
              {"fieldName": "test.offerDate",     "sheetName": "Sheet1", "cellAddress": {"row": 1, "col": 0}},
              {"fieldName": "test.currency",      "sheetName": "Sheet1", "cellAddress": {"row": 2, "col": 0}},
              {"fieldName": "test.product",       "sheetName": "Sheet1", "cellAddress": {"row": 3, "col": 0}},
              {"fieldName": "test.rateModNumber", "sheetName": "Sheet1", "cellAddress": {"row": 4, "col": 0}}
            ]}
            """;

    // JSON-шаблон: loop с заголовками.
    // cellAddress.row=1 — строка данных (вариант A: маркерная строка заменяется данными).
    // row=0 — строка заголовков (headerName).
    // При экспорте: заголовки пишутся в row=0, данные в row=1,2,3,...
    // При импорте: HEADER mode — заголовок найден в row=0, данные читаются с row=1+.
    private static final String LOOP_JSON = """
            {"entries": [
              {"fieldName": "for.dateList.origin",      "sheetName": "Sheet1", "cellAddress": {"row": 1, "col": 0}, "headerName": "ORIGIN"},
              {"fieldName": "for.dateList.destination", "sheetName": "Sheet1", "cellAddress": {"row": 1, "col": 1}, "headerName": "DESTINATION"},
              {"fieldName": "for.dateList.rate",        "sheetName": "Sheet1", "cellAddress": {"row": 1, "col": 2}, "headerName": "RATE"}
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
        assertEquals(source.getAccount(),       result.getAccount(),       "account");
        assertEquals(source.getOfferDate(),     result.getOfferDate(),     "offerDate");
        assertEquals(source.getCurrency(),      result.getCurrency(),      "currency");
        assertEquals(source.getProduct(),       result.getProduct(),       "product");
        assertEquals(source.getRateModNumber(), result.getRateModNumber(), "rateModNumber");
    }

    @Test
    void loopObjects_exportAndReimport_allRowsPresent() throws Exception {
        List<LoopDate> sourceList = List.of(
                buildLoopDate("RU", "DE", "0.05"),
                buildLoopDate("US", "FR", "0.12"),
                buildLoopDate("CN", "GB", "0.08")
        );

        byte[] xlsx = exportLoop(LOOP_JSON, sourceList);

        ExcelImportParamCore importParam = new ExcelImportParamCore();
        importParam.getParamsMap().put("for.dateList", new ImportInformation().setClazz(LoopDate.class));
        ExcelImportUtil.importExcel(importParam,
                new JsonTemplateReader(jsonStream(LOOP_JSON)),
                new ByteArrayInputStream(xlsx));

        List<Object> result = importParam.getParamsMap().get("for.dateList").getLoopLst();
        assertEquals(sourceList.size(), result.size(), "row count");

        for (int i = 0; i < sourceList.size(); i++) {
            LoopDate src = sourceList.get(i);
            LoopDate res = (LoopDate) result.get(i);
            assertEquals(src.getOrigin(),      res.getOrigin(),      "origin[" + i + "]");
            assertEquals(src.getDestination(), res.getDestination(), "destination[" + i + "]");
            assertEquals(src.getRate(),        res.getRate(),        "rate[" + i + "]");
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
        try (InputStream tmpl = resource("02_RateModNotice_tmpl.xlsx");
             InputStream data = resource("01_RateModNotice_tmpl.xlsx")) {
            ExcelImportUtil.importExcel(origParam, tmpl, data);
        }
        ExcelImport orig = (ExcelImport) origParam.getParamsMap().get("test").getLoopLst().get(0);

        // export
        ExcelExportParamCore exportParam = new ExcelExportParamCore();
        exportParam.getParamsMap().put("test", new ExportInformation().setDataList(List.of(orig)));
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (InputStream tmpl = resource("02_RateModNotice_tmpl.xlsx")) {
            ExcelExportUtil.exportExcel(exportParam, tmpl, out);
        }

        // re-import
        ExcelImportParamCore reimportParam = buildImportParam();
        try (InputStream tmpl = resource("02_RateModNotice_tmpl.xlsx")) {
            ExcelImportUtil.importExcel(reimportParam, tmpl, new ByteArrayInputStream(out.toByteArray()));
        }
        ExcelImport result = (ExcelImport) reimportParam.getParamsMap().get("test").getLoopLst().get(0);

        assertEquals(orig.getAccount(),       result.getAccount(),       "account");
        assertEquals(orig.getOfferDate(),     result.getOfferDate(),     "offerDate");
        assertEquals(orig.getCurrency(),      result.getCurrency(),      "currency");
        assertEquals(orig.getProduct(),       result.getProduct(),       "product");
        assertEquals(orig.getRateModNumber(), result.getRateModNumber(), "rateModNumber");
    }

    @Test
    void excelTemplate_withStyles_singleObject_roundtrip() throws Exception {
        // import original data
        ExcelImportParamCore origParam = buildImportParam();
        try (InputStream tmpl = resource("02_RateModNotice_tmpl.xlsx");
             InputStream data = resource("01_RateModNotice_tmpl.xlsx")) {
            ExcelImportUtil.importExcel(origParam, tmpl, data);
        }
        ExcelImport orig = (ExcelImport) origParam.getParamsMap().get("test").getLoopLst().get(0);

        // export WITH styles
        ExcelExportParamCore exportParam = new ExcelExportParamCore();
        exportParam.getParamsMap().put("test", new ExportInformation().setDataList(List.of(orig)));
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (InputStream tmpl = resource("02_RateModNotice_tmpl.xlsx")) {
            ExcelExportUtil.exportExcelWithStyles(exportParam, tmpl, out);
        }
        assertTrue(out.size() > 0, "output must not be empty");
        assertTrue(exportParam.getWarnings().isEmpty(),
                "no warnings expected, got: " + exportParam.getWarnings());

        // re-import — values must survive the styled export
        ExcelImportParamCore reimportParam = buildImportParam();
        try (InputStream tmpl = resource("02_RateModNotice_tmpl.xlsx")) {
            ExcelImportUtil.importExcel(reimportParam, tmpl, new ByteArrayInputStream(out.toByteArray()));
        }
        ExcelImport result = (ExcelImport) reimportParam.getParamsMap().get("test").getLoopLst().get(0);
        assertEquals(orig.getAccount(),       result.getAccount(),       "account");
        assertEquals(orig.getCurrency(),      result.getCurrency(),      "currency");
        assertEquals(orig.getProduct(),       result.getProduct(),       "product");
        assertEquals(orig.getRateModNumber(), result.getRateModNumber(), "rateModNumber");
    }

    @Test
    void excelTemplate_withStyles_customDepth_doesNotThrow() {
        ExcelExportParamCore exportParam = new ExcelExportParamCore();
        exportParam.getParamsMap().put("test",
                new ExportInformation().setDataList(List.of(buildSingleObject())));

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        assertDoesNotThrow(() -> {
            try (InputStream tmpl = resource("02_RateModNotice_tmpl.xlsx")) {
                ExcelExportUtil.exportExcelWithStyles(exportParam, tmpl, out, 10);
            }
        });
        assertTrue(out.size() > 0);
    }

    @Test
    void excelTemplate_loopObjects_firstRowRoundtrip() throws Exception {
        // import original data
        ExcelImportParamCore origParam = buildImportParam();
        try (InputStream tmpl = resource("02_RateModNotice_tmpl.xlsx");
             InputStream data = resource("01_RateModNotice_tmpl.xlsx")) {
            ExcelImportUtil.importExcel(origParam, tmpl, data);
        }
        List<Object> origLoop = origParam.getParamsMap().get("for.dateList").getLoopLst();
        assertFalse(origLoop.isEmpty(), "original loop must not be empty");

        // export
        ExcelExportParamCore exportParam = new ExcelExportParamCore();
        exportParam.getParamsMap().put("for.dateList", new ExportInformation().setDataList(origLoop));
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (InputStream tmpl = resource("02_RateModNotice_tmpl.xlsx")) {
            ExcelExportUtil.exportExcel(exportParam, tmpl, out);
        }

        // re-import (template has 2 sheets with loop markers → more rows than original)
        ExcelImportParamCore reimportParam = buildImportParam();
        try (InputStream tmpl = resource("02_RateModNotice_tmpl.xlsx")) {
            ExcelImportUtil.importExcel(reimportParam, tmpl, new ByteArrayInputStream(out.toByteArray()));
        }
        List<Object> result = reimportParam.getParamsMap().get("for.dateList").getLoopLst();
        assertFalse(result.isEmpty(), "re-imported loop must not be empty");

        // first row must match (comes from first sheet "Pricelist")
        LoopDate origFirst = (LoopDate) origLoop.get(0);
        LoopDate resultFirst = (LoopDate) result.get(0);
        assertEquals(origFirst.getOrigin(),      resultFirst.getOrigin(),      "origin[0]");
        assertEquals(origFirst.getDestination(), resultFirst.getDestination(), "destination[0]");
        assertEquals(origFirst.getRate(),        resultFirst.getRate(),        "rate[0]");
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
        exportParam.getParamsMap().put("for.dateList",
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
        param.getParamsMap().put("test", new ImportInformation().setClazz(ExcelImport.class));
        param.getParamsMap().put("for.dateList", new ImportInformation().setClazz(LoopDate.class));
        return param;
    }

    private InputStream resource(String filename) {
        return getClass().getResourceAsStream("/xlstemplates/" + filename);
    }

    private static ExcelImport buildSingleObject() {
        ExcelImport obj = new ExcelImport();
        obj.setAccount("ACC-001");
        obj.setOfferDate("2025-01-15");
        obj.setCurrency("USD");
        obj.setProduct("Premium");
        obj.setRateModNumber("RM-42");
        return obj;
    }

    private static LoopDate buildLoopDate(String origin, String destination, String rate) {
        LoopDate ld = new LoopDate();
        ld.setOrigin(origin);
        ld.setDestination(destination);
        ld.setRate(rate);
        return ld;
    }
}
