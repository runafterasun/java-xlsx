package ru.objectsfill.model;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * Информация об одном объекте импорта: целевой класс, карта маркеров из шаблона,
 * список наполненных экземпляров и перечень обязательных полей.
 */
public class ImportInformation {

    /** Целевой класс, экземпляры которого создаются при импорте. */
    private Class<?> clazz;

    /** Список наполненных объектов, полученных в результате импорта. */
    private List<Object> loopLst = new ArrayList<>();

    /** Произвольный тег для идентификации блока. */
    private String tag;

    /**
     * Карта маркеров, сгруппированных по имени листа.
     * Заполняется после анализа шаблона.
     */
    private Map<String, List<FieldParam>> paramMap;

    /** Список обязательных полей, которые должны быть найдены при импорте. */
    private List<String> requiredFields = new ArrayList<>();

    /**
     * Возвращает целевой класс для создания экземпляров при импорте.
     *
     * @return целевой класс
     */
    public Class<?> getClazz() {
        return clazz;
    }

    /**
     * Устанавливает целевой класс.
     *
     * @param clazz целевой класс
     * @return текущий экземпляр для цепочки вызовов
     */
    public ImportInformation setClazz(Class<?> clazz) {
        this.clazz = clazz;
        return this;
    }

    /**
     * Возвращает список наполненных объектов.
     *
     * @return список результатов импорта
     */
    public List<Object> getLoopLst() {
        return loopLst;
    }

    /**
     * Устанавливает список наполненных объектов.
     *
     * @param loopLst новый список
     * @return текущий экземпляр для цепочки вызовов
     */
    public ImportInformation setLoopLst(List<Object> loopLst) {
        this.loopLst = loopLst;
        return this;
    }

    /**
     * Возвращает произвольный тег блока.
     *
     * @return тег
     */
    public String getTag() {
        return tag;
    }

    /**
     * Устанавливает произвольный тег блока.
     *
     * @param tag тег
     * @return текущий экземпляр для цепочки вызовов
     */
    public ImportInformation setTag(String tag) {
        this.tag = tag;
        return this;
    }

    /**
     * Возвращает карту маркеров, сгруппированных по имени листа.
     *
     * @return карта маркеров
     */
    public Map<String, List<FieldParam>> getParamMap() {
        return paramMap;
    }

    /**
     * Устанавливает карту маркеров.
     *
     * @param paramMap карта маркеров, ключ — имя листа
     * @return текущий экземпляр для цепочки вызовов
     */
    public ImportInformation setParamMap(Map<String, List<FieldParam>> paramMap) {
        this.paramMap = paramMap;
        return this;
    }

    /**
     * Возвращает список обязательных полей.
     *
     * @return список имён обязательных полей
     */
    public List<String> getRequiredFields() {
        return requiredFields;
    }

    /**
     * Устанавливает список обязательных полей.
     *
     * @param requiredFields список имён обязательных полей
     * @return текущий экземпляр для цепочки вызовов
     */
    public ImportInformation setRequiredFields(List<String> requiredFields) {
        this.requiredFields = requiredFields;
        return this;
    }
}
