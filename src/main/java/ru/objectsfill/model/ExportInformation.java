package ru.objectsfill.model;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExportInformation {

    private List<Object> dataList = new ArrayList<>();

    private Map<String, List<FieldParam>> paramMap = new HashMap<>();

    public List<Object> getDataList() {
        return dataList;
    }

    public ExportInformation setDataList(List<Object> dataList) {
        this.dataList = dataList;
        return this;
    }

    public Map<String, List<FieldParam>> getParamMap() {
        return paramMap;
    }

    public ExportInformation setParamMap(Map<String, List<FieldParam>> paramMap) {
        this.paramMap = paramMap;
        return this;
    }
}
