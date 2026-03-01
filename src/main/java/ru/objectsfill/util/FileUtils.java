package ru.objectsfill.util;

import ru.objectsfill.exception.ExcelImportException;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.Arrays;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

public class FileUtils {

    private FileUtils() {}

    private static final Map<Class<?>, Field[]> declaredFieldsCache = new ConcurrentHashMap<>(256);


    /**
     * @param clazz с которого получаем все поля
     * @return список полей класса для наполнения
     */
    private static Field[] getDeclaredFields(Class<?> clazz) {
        if (clazz == null)
            throw new IllegalArgumentException("Class must not be null");
        var result = declaredFieldsCache.get(clazz);
        if (result == null) {
            try {
                result = clazz.getDeclaredFields();
                declaredFieldsCache.put(clazz, (result.length == 0 ? new Field[0] : result));
            } catch (Exception ex) {
                throw new IllegalStateException("Failed to introspect Class [" + clazz.getName() +
                        "] from ClassLoader [" + clazz.getClassLoader() + "]", ex);
            }
        }
        return result;
    }

    /**
     * Записывает значение {@code val} в поле {@code fieldName} объекта {@code object}
     * через соответствующий setter. Если поле с таким именем отсутствует — ничего не делает.
     *
     * @param object    объект, который наполняем
     * @param val       значение для записи
     * @param fieldName имя поля целевого объекта
     * @throws ExcelImportException если setter не найден или его вызов завершился ошибкой
     */
    public static void fillObject(Object object, Object val, String fieldName) {
        var declaredFields = object.getClass().getDeclaredFields();
        Arrays.stream(declaredFields).forEach(field -> {
            if (field.getName().equals(fieldName)) {
                try {
                    Method method = object.getClass().getMethod("set" + field.getName().substring(0, 1).toUpperCase() + field.getName().substring(1), field.getType());
                    method.invoke(object, val);
                } catch (Exception ex) {
                    throw new ExcelImportException(
                            "Failed to set field '" + fieldName + "' on " + object.getClass().getName(), ex);
                }
            }
        });
    }
}
