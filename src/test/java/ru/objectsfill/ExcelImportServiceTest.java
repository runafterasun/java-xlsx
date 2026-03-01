package ru.objectsfill;

import org.junit.jupiter.api.Test;
import ru.objectsfill.dto.ExcelImport;
import ru.objectsfill.dto.LoopDate;
import ru.objectsfill.model.ExcelImportParamCore;
import ru.objectsfill.model.ImportInformation;

import java.io.InputStream;

import static org.junit.jupiter.api.Assertions.*;

class ExcelImportServiceTest {

    @Test
    void sheet0_singleObjectIsFilled() throws Exception {
        var importParam = new ExcelImportParamCore();

        importParam.getParamsMap().put("test", new ImportInformation().setClazz(ExcelImport.class));
        importParam.getParamsMap().put("for.dateList", new ImportInformation().setClazz(LoopDate.class));
        importParam.getParamsMap().put("for.testSheet", new ImportInformation().setClazz(LoopDate.class));

        runImport(importParam);

        ExcelImport excelImport = (ExcelImport) importParam.getParamsMap().get("test").getLoopLst().get(0);
        assertNotNull(excelImport.getAccount(), "account");
        assertNotNull(excelImport.getOfferDate(), "offerDate");
        assertNotNull(excelImport.getCurrency(), "currency");
        assertNotNull(excelImport.getProduct(), "product");
        assertNotNull(excelImport.getRateModNumber(), "rateModNumber");
    }

    @Test
    void sheet0_loopListIsPopulated() throws Exception {
        var importParam = new ExcelImportParamCore();

        importParam.getParamsMap().put("test", new ImportInformation().setClazz(ExcelImport.class));
        importParam.getParamsMap().put("for.dateList", new ImportInformation().setClazz(LoopDate.class));

        runImport(importParam);

        assertFalse(importParam.getParamsMap().get("for.dateList").getLoopLst().isEmpty(), "loop list should not be empty");
    }

    private void runImport(ExcelImportParamCore param) throws Exception {
        try (InputStream tmpl = resource("02_RateModNotice_tmpl.xlsx");
             InputStream data = resource("01_RateModNotice_tmpl.xlsx")) {
            ExcelImportUtil.importExcel(param, tmpl, data);
        }
    }

    private InputStream resource(String filename) {
        return getClass().getResourceAsStream("/xlstemplates/" + filename);
    }
}
