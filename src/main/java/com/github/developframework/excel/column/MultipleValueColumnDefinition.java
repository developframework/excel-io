package com.github.developframework.excel.column;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Workbook;

import java.lang.reflect.Field;
import java.util.Arrays;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.stream.Stream;

/**
 * 多值列定义
 *
 * @author qiushui on 2019-01-16.
 */
public class MultipleValueColumnDefinition extends BasicColumnDefinition {

    private String separator;

    public MultipleValueColumnDefinition(Workbook workbook, String field) {
        this(workbook, field, null);
    }

    public MultipleValueColumnDefinition(Workbook workbook, String field, String separator) {
        super(workbook, field);
        this.cellType = CellType.STRING;
        this.cellStyle.setWrapText(true);
        this.separator = separator == null ? "\n" : separator;
    }

    @Override
    public void dealFillData(Cell cell, Object value) {
        String content;
        if (value.getClass().isArray()) {
            content = StringUtils.join((Object[]) value, separator);
        } else if (value instanceof List) {
            content = StringUtils.join((List) value, separator);
        } else if (value instanceof Set) {
            content = StringUtils.join((Set) value, separator);
        } else {
            throw new IllegalArgumentException();
        }
        cell.setCellValue(content);
    }

    @Override
    public void dealReadData(Cell cell, Object instance) {
        String content = cell.getStringCellValue();
        String[] parts = Stream.of(content.split(separator)).map(item -> readColumnValueConverter.map(converter -> converter.convert(instance, item)).orElse(item)).toArray(String[]::new);
        Class<?> instanceClass = instance.getClass();
        Field field = FieldUtils.getField(instanceClass, fieldName, true);
        try {
            if (field.getType().isArray()) {
                FieldUtils.writeDeclaredField(instance, fieldName, parts, true);
            } else if (field.getType() == List.class) {
                List list = Arrays.asList(parts);
                FieldUtils.writeDeclaredField(instance, fieldName, list, true);
            } else if (field.getType() == Set.class) {
                Set<String> set = new HashSet<>(Arrays.asList(parts));
                FieldUtils.writeDeclaredField(instance, fieldName, set, true);
            } else {
                throw new IllegalArgumentException();
            }
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
    }
}
