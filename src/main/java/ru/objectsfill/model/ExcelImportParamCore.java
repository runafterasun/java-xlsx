package ru.objectsfill.model;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Корневой контейнер параметров импорта.
 * Хранит карту, где ключ — идентификатор объекта (например, {@code "test"} или {@code "for.dateList"}),
 * а значение — {@link ImportInformation} с целевым классом и результатами импорта.
 */
public class ExcelImportParamCore {

    /** Список параметров, которые надо найти при чтении файла шаблона. */
    private Map<String, ImportInformation> paramsMap = new HashMap<>();

    /** Предупреждения, собранные в ходе импорта (например, поле не найдено в объекте). */
    private final List<String> warnings = new ArrayList<>();

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

    /**
     * Возвращает список предупреждений, собранных в ходе импорта.
     *
     * @return изменяемый список предупреждений
     */
    public List<String> getWarnings() {
        return warnings;
    }
}
