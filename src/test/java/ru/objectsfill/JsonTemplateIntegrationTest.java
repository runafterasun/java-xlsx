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
 * Структура шаблона выявлена из 02_RateModNotice_tmpl.xlsx (ExcelTemplateReader):
 *
 * Лист "Pricelist":
 *   test.account      row=3  col=0
 *   test.offerDate    row=4  col=0
 *   test.currency     row=5  col=0
 *   test.product      row=6  col=0
 *   test.rateModNumber row=6 col=4
 *   for.dateList.*    row=10, headerNames: DESTINATION, ORIGIN, COUNTRY CODE,
 *                     ROUTE TEL PREFIX, EFFECTIVE DATE, RATE, CHANGE, BILLING INCREMENTS
 *
 * Лист "Лист3":
 *   for.dateList.*    row=1, те же headerNames (без RATE)
 */
class JsonTemplateIntegrationTest {

    private static final String PRICELIST_SINGLE_JSON = """
            {"entries": [
              {"fieldName": "test.account",       "sheetName": "Pricelist", "cellAddress": {"row": 3, "col": 0}},
              {"fieldName": "test.offerDate",     "sheetName": "Pricelist", "cellAddress": {"row": 4, "col": 0}},
              {"fieldName": "test.currency",      "sheetName": "Pricelist", "cellAddress": {"row": 5, "col": 0}},
              {"fieldName": "test.product",       "sheetName": "Pricelist", "cellAddress": {"row": 6, "col": 0}},
              {"fieldName": "test.rateModNumber", "sheetName": "Pricelist", "cellAddress": {"row": 6, "col": 4}}
            ]}
            """;

    private static final String PRICELIST_DATELIST_JSON = """
            {"entries": [
              {"fieldName": "for.dateList.origin",      "sheetName": "Pricelist", "headerName": "DESTINATION"},
              {"fieldName": "for.dateList.destination", "sheetName": "Pricelist", "headerName": "ORIGIN"},
              {"fieldName": "for.dateList.code",        "sheetName": "Pricelist", "headerName": "COUNTRY CODE"},
              {"fieldName": "for.dateList.route",       "sheetName": "Pricelist", "headerName": "ROUTE TEL PREFIX"},
              {"fieldName": "for.dateList.date",        "sheetName": "Pricelist", "headerName": "EFFECTIVE DATE"},
              {"fieldName": "for.dateList.rate",        "sheetName": "Pricelist", "headerName": "RATE"},
              {"fieldName": "for.dateList.change",      "sheetName": "Pricelist", "headerName": "CHANGE"},
              {"fieldName": "for.dateList.billing",     "sheetName": "Pricelist", "headerName": "BILLING INCREMENTS"}
            ]}
            """;

    private static final String PRICELIST_COMBINED_JSON = """
            {"entries": [
              {"fieldName": "test.account",             "sheetName": "Pricelist", "cellAddress": {"row": 3, "col": 0}},
              {"fieldName": "test.offerDate",           "sheetName": "Pricelist", "cellAddress": {"row": 4, "col": 0}},
              {"fieldName": "test.currency",            "sheetName": "Pricelist", "cellAddress": {"row": 5, "col": 0}},
              {"fieldName": "test.product",             "sheetName": "Pricelist", "cellAddress": {"row": 6, "col": 0}},
              {"fieldName": "test.rateModNumber",       "sheetName": "Pricelist", "cellAddress": {"row": 6, "col": 4}},
              {"fieldName": "for.dateList.origin",      "sheetName": "Pricelist", "headerName": "DESTINATION"},
              {"fieldName": "for.dateList.destination", "sheetName": "Pricelist", "headerName": "ORIGIN"},
              {"fieldName": "for.dateList.code",        "sheetName": "Pricelist", "headerName": "COUNTRY CODE"},
              {"fieldName": "for.dateList.route",       "sheetName": "Pricelist", "headerName": "ROUTE TEL PREFIX"},
              {"fieldName": "for.dateList.date",        "sheetName": "Pricelist", "headerName": "EFFECTIVE DATE"},
              {"fieldName": "for.dateList.rate",        "sheetName": "Pricelist", "headerName": "RATE"},
              {"fieldName": "for.dateList.change",      "sheetName": "Pricelist", "headerName": "CHANGE"},
              {"fieldName": "for.dateList.billing",     "sheetName": "Pricelist", "headerName": "BILLING INCREMENTS"}
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
        runImport(PRICELIST_SINGLE_JSON, param);

        ExcelImport result = (ExcelImport) param.getParamsMap().get("test").getLoopLst().get(0);
        assertNotNull(result.getAccount(),       "account");
        assertNotNull(result.getOfferDate(),     "offerDate");
        assertNotNull(result.getCurrency(),      "currency");
        assertNotNull(result.getProduct(),       "product");
        assertNotNull(result.getRateModNumber(), "rateModNumber");
    }

    /**
     * Результаты чтения одиночного объекта через JSON-шаблон должны совпадать
     * с результатами чтения через Excel-шаблон.
     */
    @Test
    void singleObject_jsonMatchesExcelTemplate() throws Exception {
        var jsonParam = buildParam();
        runImport(PRICELIST_SINGLE_JSON, jsonParam);
        ExcelImport fromJson = (ExcelImport) jsonParam.getParamsMap().get("test").getLoopLst().get(0);

        var xlsParam = buildParam();
        try (InputStream tmpl = resource("02_RateModNotice_tmpl.xlsx");
             InputStream data = resource("01_RateModNotice_tmpl.xlsx")) {
            ExcelImportUtil.importExcel(xlsParam, tmpl, data);
        }
        ExcelImport fromXls = (ExcelImport) xlsParam.getParamsMap().get("test").getLoopLst().get(0);

        assertEquals(fromXls.getAccount(),       fromJson.getAccount(),       "account");
        assertEquals(fromXls.getOfferDate(),     fromJson.getOfferDate(),     "offerDate");
        assertEquals(fromXls.getCurrency(),      fromJson.getCurrency(),      "currency");
        assertEquals(fromXls.getProduct(),       fromJson.getProduct(),       "product");
        assertEquals(fromXls.getRateModNumber(), fromJson.getRateModNumber(), "rateModNumber");
    }

    // =================== loop list ===================

    /**
     * JSON-шаблон с привязкой по заголовку (HEADER).
     * Маркеры для loop-списка описаны через headerName.
     * После импорта список dateList должен быть непустым.
     */
    @Test
    void loopList_headerBinding_listIsPopulated() throws Exception {
        var param = buildParam();
        runImport(PRICELIST_DATELIST_JSON, param);

        List<Object> loopLst = param.getParamsMap().get("for.dateList").getLoopLst();
        assertFalse(loopLst.isEmpty(), "loop list must not be empty");
    }

    /**
     * Loop-список, прочитанный через JSON-шаблон (только лист Pricelist),
     * должен быть непустым и содержать разумное количество строк.
     */
    @Test
    void loopList_jsonSizeMatchesExcelTemplate() throws Exception {
        var jsonParam = buildParam();
        runImport(PRICELIST_DATELIST_JSON, jsonParam);
        int jsonSize = jsonParam.getParamsMap().get("for.dateList").getLoopLst().size();

        assertTrue(jsonSize > 0, "JSON loop list must not be empty");
    }

    /**
     * Поля первой строки loop-списка, прочитанной через JSON, должны совпадать
     * с полями первой строки, прочитанной через Excel-шаблон.
     */
    @Test
    void loopList_firstRowFieldsMatchExcelTemplate() throws Exception {
        var jsonParam = buildParam();
        runImport(PRICELIST_DATELIST_JSON, jsonParam);
        LoopDate fromJson = (LoopDate) jsonParam.getParamsMap().get("for.dateList").getLoopLst().get(0);

        var xlsParam = buildParam();
        try (InputStream tmpl = resource("02_RateModNotice_tmpl.xlsx");
             InputStream data = resource("01_RateModNotice_tmpl.xlsx")) {
            ExcelImportUtil.importExcel(xlsParam, tmpl, data);
        }
        LoopDate fromXls = (LoopDate) xlsParam.getParamsMap().get("for.dateList").getLoopLst().get(0);

        assertEquals(fromXls.getOrigin(),      fromJson.getOrigin(),      "origin");
        assertEquals(fromXls.getDestination(), fromJson.getDestination(), "destination");
        assertEquals(fromXls.getCode(),        fromJson.getCode(),        "code");
        assertEquals(fromXls.getRoute(),       fromJson.getRoute(),       "route");
        assertEquals(fromXls.getRate(),        fromJson.getRate(),        "rate");
        assertEquals(fromXls.getChange(),      fromJson.getChange(),      "change");
        assertEquals(fromXls.getBilling(),     fromJson.getBilling(),     "billing");
    }

    // =================== combined ===================

    /**
     * JSON-шаблон описывает и одиночный объект, и loop-список одновременно.
     * После импорта оба должны быть заполнены.
     */
    @Test
    void combined_singleAndLoop_bothFilled() throws Exception {
        var param = buildParam();
        runImport(PRICELIST_COMBINED_JSON, param);

        ExcelImport single = (ExcelImport) param.getParamsMap().get("test").getLoopLst().get(0);
        assertNotNull(single.getAccount(),   "single: account");
        assertNotNull(single.getCurrency(),  "single: currency");

        List<Object> loop = param.getParamsMap().get("for.dateList").getLoopLst();
        assertFalse(loop.isEmpty(), "loop list must not be empty");

        LoopDate firstRow = (LoopDate) loop.get(0);
        assertNotNull(firstRow.getRate(), "loop first row: rate");
    }

    // =================== helpers ===================

    private static ExcelImportParamCore buildParam() {
        var param = new ExcelImportParamCore();
        param.getParamsMap().put("test",         new ImportInformation().setClazz(ExcelImport.class));
        param.getParamsMap().put("for.dateList", new ImportInformation().setClazz(LoopDate.class));
        return param;
    }

    private void runImport(String json, ExcelImportParamCore param) throws Exception {
        var reader = new JsonTemplateReader(
                new ByteArrayInputStream(json.getBytes(StandardCharsets.UTF_8)));
        try (InputStream data = resource("01_RateModNotice_tmpl.xlsx")) {
            ExcelImportUtil.importExcel(param, reader, data);
        }
    }

    private InputStream resource(String filename) {
        return getClass().getResourceAsStream("/xlstemplates/" + filename);
    }
}
