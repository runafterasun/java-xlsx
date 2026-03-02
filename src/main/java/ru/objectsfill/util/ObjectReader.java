package ru.objectsfill.util;

import ru.objectsfill.exception.ExcelExportException;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.List;

public class ObjectReader {

    private ObjectReader() {}

    /**
     * Читает значение поля объекта через getter-метод.
     * Если поле не найдено — добавляет предупреждение и возвращает {@code null}.
     *
     * @param object    объект, из которого читается поле
     * @param fieldName имя поля
     * @param warnings  список предупреждений
     * @return значение поля или {@code null}, если поле отсутствует
     */
    public static Object readField(Object object, String fieldName, List<String> warnings) {
        for (Field field : object.getClass().getDeclaredFields()) {
            if (field.getName().equals(fieldName)) {
                try {
                    String getter = "get"
                            + field.getName().substring(0, 1).toUpperCase()
                            + field.getName().substring(1);
                    Method method = object.getClass().getMethod(getter);
                    return method.invoke(object);
                } catch (Exception ex) {
                    throw new ExcelExportException(
                            "Failed to read field '" + fieldName + "' on " + object.getClass().getName(), ex);
                }
            }
        }
        warnings.add("Field '" + fieldName + "' not found on " + object.getClass().getSimpleName());
        return null;
    }
}
