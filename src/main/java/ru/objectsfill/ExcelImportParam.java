package ru.objectsfill;

import lombok.Getter;
import lombok.Setter;
import lombok.experimental.Accessors;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Getter
@Setter
@Accessors(chain = true)
public class ExcelImportParam<T> {

    private Map<String, Object> paramsMap = new HashMap<>();
    private List<T> loopLst = new ArrayList<>();
    private Integer wbSheetNumber;
}
