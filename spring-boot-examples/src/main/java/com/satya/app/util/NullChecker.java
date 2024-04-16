package com.satya.app.util;

import java.lang.reflect.Field;
import java.util.Arrays;
import java.util.Objects;

import org.apache.commons.lang3.StringUtils;

public class NullChecker {

    public static boolean allNull(Object target) {
        return Arrays.stream(target.getClass()
          .getDeclaredFields())
          .peek(f -> f.setAccessible(true))
          .map(f -> getFieldValue(f, target))
          .allMatch(Objects::isNull);
    }
    
    public static boolean allNullExcept(Object target, String fieldName) {
        return Arrays.stream(target.getClass()
          .getDeclaredFields())
          .filter(f -> !StringUtils.equals(fieldName, f.getName()))
          .peek(f -> f.setAccessible(true))
          .map(f -> getFieldValue(f, target))
          .allMatch(Objects::isNull);
    }

    private static Object getFieldValue(Field field, Object target) {
        try {
            return field.get(target);
        } catch (IllegalAccessException e) {
            throw new RuntimeException(e);
        }
    }
}
