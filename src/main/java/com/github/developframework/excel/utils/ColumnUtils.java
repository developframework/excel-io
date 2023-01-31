package com.github.developframework.excel.utils;

/**
 * @author qiushui on 2023-01-31.
 */
public abstract class ColumnUtils {

    public static String getColumnNameByIndex(short index) {
        int a = index / 26;
        if (a == 0) {
            return String.valueOf((char) ('A' + index % 26));
        } else {
            return (char) ('A' + a - 1) + String.valueOf((char) ('A' + index % 26));
        }
    }

    public static short getColumnIndexByName(String name) {
        final int length = name.length();
        if (length > 2) {
            throw new IllegalArgumentException("错误的列名" + name);
        }
        if (length == 1) {
            return (short) (name.charAt(0) - 'A');
        } else {
            return (short) ((name.charAt(0) - 'A') * 26 + (name.charAt(1) - 'A'));
        }
    }
}
