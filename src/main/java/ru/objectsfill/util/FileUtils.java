package ru.objectsfill.util;

import ru.objectsfill.exception.ExcelImportException;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.List;

public class FileUtils {

    private FileUtils() {}

    /**
     * Записывает значение {@code val} в поле {@code fieldName} объекта {@code object}
     * через соответствующий setter. Если поле с таким именем отсутствует — добавляет
     * предупреждение в {@code warnings}.
     *
     * @param object    объект, который наполняем
     * @param val       значение для записи
     * @param fieldName имя поля целевого объекта
     * @param warnings  список для сбора предупреждений
     * @throws ExcelImportException если setter не найден или его вызов завершился ошибкой
     */
    public static void fillObject(Object object, Object val, String fieldName, List<String> warnings) {
        Field[] declaredFields = object.getClass().getDeclaredFields();
        for (Field field : declaredFields) {
            if (field.getName().equals(fieldName)) {
                try {
                    Method method = object.getClass().getMethod(
                            "set" + field.getName().substring(0, 1).toUpperCase() + field.getName().substring(1),
                            field.getType());
                    method.invoke(object, val);
                } catch (Exception ex) {
                    throw new ExcelImportException(
                            "Failed to set field '" + fieldName + "' on " + object.getClass().getName(), ex);
                }
                return;
            }
        }
        warnings.add("Field '" + fieldName + "' not found on " + object.getClass().getSimpleName());
    }
}
