package com.github.developframework.excel;

import java.math.BigDecimal;
import java.time.*;
import java.time.format.DateTimeFormatter;
import java.util.Date;

/**
 * @author qiushui on 2022-06-30.
 */
public class ValueConvertUtils {

    /**
     * 字符串转换
     */
    public static Object stringConvert(String value, Class<?> targetClass) {
        if (value == null) {
            return null;
        } else if (targetClass == String.class) {
            return value;
        } else if (targetClass == Integer.class || targetClass == Integer.TYPE) {
            return new BigDecimal(value).intValue();
        } else if (targetClass == Long.class || targetClass == Long.TYPE) {
            return new BigDecimal(value).longValue();
        } else if (targetClass == BigDecimal.class) {
            return new BigDecimal(value);
        } else if (targetClass == Short.class || targetClass == Short.TYPE) {
            return Short.parseShort(value);
        } else if (targetClass == Float.class || targetClass == Float.TYPE) {
            return Float.parseFloat(value);
        } else if (targetClass == Double.class || targetClass == Double.TYPE) {
            return Double.parseDouble(value);
        } else if (targetClass == Boolean.class || targetClass == Boolean.TYPE) {
            return Boolean.parseBoolean(value);
        } else if (targetClass == java.util.Date.class) {
            return Date.from((LocalDateTime.parse(value, DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"))).atZone(ZoneId.systemDefault()).toInstant());
        } else if (targetClass == LocalDateTime.class) {
            return LocalDateTime.parse(value, DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));
        } else if (targetClass == ZonedDateTime.class) {
            return ZonedDateTime.parse(value, DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));
        } else if (targetClass == LocalDate.class) {
            return LocalDate.parse(value);
        } else if (targetClass == LocalTime.class) {
            return LocalTime.parse(value);
        } else if (targetClass.isEnum()) {
            return Enum.valueOf((Class<Enum>) targetClass, value);
        } else {
            return value;
        }
    }

    public static Object doubleConvert(double value, Class<?> targetClass) {
        if (targetClass == Double.class || targetClass == Double.TYPE) {
            return value;
        } else if (targetClass == String.class) {
            return Double.toString(value);
        } else if (targetClass == Integer.class || targetClass == Integer.TYPE) {
            return (int) value;
        } else if (targetClass == Long.class || targetClass == Long.TYPE) {
            return (long) value;
        } else if (targetClass == Short.class || targetClass == Short.TYPE) {
            return (short) value;
        } else if (targetClass == Float.class || targetClass == Float.TYPE) {
            return (float) value;
        } else if (targetClass == Boolean.class || targetClass == Boolean.TYPE) {
            return value > 0;
        } else {
            return value;
        }
    }

    public static Object booleanConvert(boolean value, Class<?> targetClass) {
        if (targetClass == Boolean.class || targetClass == Boolean.TYPE) {
            return value;
        } else if (targetClass == String.class) {
            return Boolean.toString(value);
        } else if (targetClass == Integer.class || targetClass == Integer.TYPE) {
            return value ? 1 : 0;
        } else if (targetClass == Long.class || targetClass == Long.TYPE) {
            return value ? 1L : 0L;
        } else if (targetClass == Short.class || targetClass == Short.TYPE) {
            return (short) (value ? 1 : 0);
        } else if (targetClass == Float.class || targetClass == Float.TYPE) {
            return value ? 1f : 0f;
        } else if (targetClass == Double.class || targetClass == Double.TYPE) {
            return value ? 1d : 0d;
        } else {
            return value;
        }
    }

    public static Object dateConvert(Date value, Class<?> targetClass) {
        if (targetClass == Date.class) {
            return value;
        } else if (targetClass == LocalDateTime.class) {
            return LocalDateTime.ofInstant(value.toInstant(), ZoneId.systemDefault());
        } else if (targetClass == ZonedDateTime.class) {
            return ZonedDateTime.ofInstant(value.toInstant(), ZoneId.systemDefault());
        } else if (targetClass == LocalDate.class) {
            return LocalDateTime.ofInstant(value.toInstant(), ZoneId.systemDefault()).toLocalDate();
        } else if (targetClass == LocalTime.class) {
            return LocalDateTime.ofInstant(value.toInstant(), ZoneId.systemDefault()).toLocalTime();
        } else {
            return value;
        }
    }
}
