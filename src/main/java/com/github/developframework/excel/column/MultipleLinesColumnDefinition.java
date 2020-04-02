package com.github.developframework.excel.column;

import com.github.developframework.excel.ColumnDefinition;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Workbook;

import java.lang.reflect.Array;
import java.util.Collection;
import java.util.List;
import java.util.Set;
import java.util.function.Function;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * 多行值列
 */
public class MultipleLinesColumnDefinition extends ColumnDefinition<String> {

    private static final String DEFAULT_SEPARATOR = "\n";

    private Function<?, String> writeMappingFunction;

    private Function<String, ?> readMappingFunction;

    public MultipleLinesColumnDefinition(Workbook workbook, String field, String header) {
        super(workbook, field, header);
    }

    @Override
    protected CellType getColumnCellType() {
        return CellType.STRING;
    }

    @Override
    protected void setCellValue(Cell cell, String convertValue) {
        cell.getCellStyle().setWrapText(true);
        cell.setCellValue(convertValue);
    }

    @Override
    @SuppressWarnings({"unchecked", "rawtypes"})
    protected String writeConvertValue(Object entity, Object fieldValue) {
        if (fieldValue == null) {
            return null;
        }
        Stream stream;
        if (fieldValue.getClass().isArray()) {
            stream = Stream.of((Object[]) fieldValue);
        } else if (fieldValue instanceof Collection) {
            stream = ((Collection) fieldValue).stream();
        } else {
            stream = Stream.of(fieldValue);
        }
        if (writeMappingFunction != null) {
            stream = stream.map(writeMappingFunction);
        } else {
            stream = stream.map(Object::toString);
        }
        return (String) stream.collect(Collectors.joining(DEFAULT_SEPARATOR));
    }

    @Override
    protected String getCellValue(Cell cell) {
        String cellValue = dataFormatter.formatCellValue(cell);
        return cellValue == null ? null : cellValue.trim();
    }

    @Override
    @SuppressWarnings({"unchecked", "rawtypes"})
    protected <T> Object readConvertValue(Object entity, String cellValue, Class<T> fieldClass) {
        if (cellValue == null) {
            return null;
        }
        Stream stream = Stream.of(cellValue.split(DEFAULT_SEPARATOR));
        if (readMappingFunction != null) {
            stream = stream.map(readMappingFunction);
        }
        if (fieldClass.isArray()) {
            List list = (List) stream.collect(Collectors.toList());
            Object[] array = (Object[]) Array.newInstance(fieldClass.getComponentType(), list.size());
            for (int i = 0; i < list.size(); i++) {
                array[i] = list.get(i);
            }
            return array;
        } else if (List.class.isAssignableFrom(fieldClass)) {
            return stream.collect(Collectors.toList());
        } else if (Set.class.isAssignableFrom(fieldClass)) {
            return stream.collect(Collectors.toSet());
        } else {
            return null;
        }
    }

    public MultipleLinesColumnDefinition writeMap(Function<?, String> writeMappingFunction) {
        this.writeMappingFunction = writeMappingFunction;
        return this;
    }

    public MultipleLinesColumnDefinition readMap(Function<String, ?> readMappingFunction) {
        this.readMappingFunction = readMappingFunction;
        return this;
    }
}
