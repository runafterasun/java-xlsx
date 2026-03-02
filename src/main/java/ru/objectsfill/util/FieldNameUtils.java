package ru.objectsfill.util;

public class FieldNameUtils {

    private FieldNameUtils() {}

    public static String getFieldName(String markerKey) {
        return markerKey.substring(markerKey.lastIndexOf(".") + 1);
    }

    public static String getKey(String markerKey) {
        int dot = markerKey.lastIndexOf(".");
        return dot != -1 ? markerKey.substring(0, dot) : markerKey;
    }
}
