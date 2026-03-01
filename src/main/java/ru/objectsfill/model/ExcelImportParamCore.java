package ru.objectsfill.model;

import java.util.HashMap;
import java.util.Map;

/**
 * Корневой контейнер параметров импорта.
 * Хранит карту, где ключ — идентификатор объекта (например, {@code "test"} или {@code "for.dateList"}),
 * а значение — {@link ImportInformation} с целевым классом и результатами импорта.
 */
public class ExcelImportParamCore {

    /** Список параметров, которые надо найти при чтении файла шаблона. */
    private Map<String, ImportInformation> paramsMap = new HashMap<>();

    /**
     * Возвращает карту параметров импорта.
     *
     * @return карта, где ключ — идентификатор объекта, значение — {@link ImportInformation}
     */
    public Map<String, ImportInformation> getParamsMap() {
        return paramsMap;
    }

    /**
     * Устанавливает карту параметров импорта.
     *
     * @param paramsMap новая карта параметров
     * @return текущий экземпляр для цепочки вызовов
     */
    public ExcelImportParamCore setParamsMap(Map<String, ImportInformation> paramsMap) {
        this.paramsMap = paramsMap;
        return this;
    }
}
