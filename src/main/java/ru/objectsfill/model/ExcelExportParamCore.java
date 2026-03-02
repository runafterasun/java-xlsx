package ru.objectsfill.model;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelExportParamCore {

    private Map<String, ExportInformation> paramsMap = new HashMap<>();

    private final List<String> warnings = new ArrayList<>();

    public Map<String, ExportInformation> getParamsMap() {
        return paramsMap;
    }

    public ExcelExportParamCore setParamsMap(Map<String, ExportInformation> paramsMap) {
        this.paramsMap = paramsMap;
        return this;
    }

    public List<String> getWarnings() {
        return warnings;
    }
}
