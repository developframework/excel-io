package com.github.developframework.excel.column;

import com.github.developframework.excel.ColumnDefinition;
import com.github.developframework.excel.ColumnValueConverter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Workbook;

import java.math.BigDecimal;

/**
 * 字符串列
 *
 * @author qiushui on 2019-05-18.
 */
@SuppressWarnings("rawtypes")
public class StringColumnDefinition extends ColumnDefinition<String> {

    private StringWriteColumnValueConverter writeColumnValueConverter;

    private StringReadColumnValueConverter readColumnValueConverter;

    protected StringColumnDefinition(Workbook workbook, String field, String header) {
        super(workbook, field, header);
    }

    @Override
    protected CellType getColumnCellType() {
        return CellType.STRING;
    }

    @Override
    protected void setCellValue(Cell cell, String convertValue) {
        cell.setCellValue(convertValue);
    }

    @Override
    protected String getCellValue(Cell cell) {
        String cellValue = dataFormatter.formatCellValue(cell);
        return cellValue == null ? null : cellValue.trim();
    }

    @SuppressWarnings("unchecked")
    @Override
    protected <T> Object readConvertValue(Object entity, String cellValue, Class<T> fieldClass) {
        Object convertValue;
        if (readColumnValueConverter != null) {
            convertValue = readColumnValueConverter.convert(entity, cellValue);
        } else {
            convertValue = cellValue;
        }
        if (convertValue == null) {
            return null;
        } else if (fieldClass == convertValue.getClass()) {
            return convertValue;
        } else if (fieldClass == String.class) {
            return convertValue;
        } else if (fieldClass == Integer.class || fieldClass == int.class) {
            return Integer.valueOf(convertValue.toString());
        } else if (fieldClass == Long.class || fieldClass == long.class) {
            return Long.valueOf(convertValue.toString());
        } else if (fieldClass == Boolean.class || fieldClass == boolean.class) {
            return Boolean.valueOf(convertValue.toString());
        } else if (fieldClass == BigDecimal.class) {
            return new BigDecimal(convertValue.toString());
        } else if (fieldClass == Float.class || fieldClass == float.class) {
            return Float.valueOf(convertValue.toString());
        } else if (fieldClass == Double.class || fieldClass == double.class) {
            return Double.valueOf(convertValue.toString());
        } else {
            throw new IllegalArgumentException("can not convert from \"java.lang.String\" to \"" + fieldClass.getName() + "\"");
        }
    }

    @SuppressWarnings("unchecked")
    @Override
    protected String writeConvertValue(Object entity, Object fieldValue) {
        if (writeColumnValueConverter == null) {
            return fieldValue == null ? null : fieldValue.toString();
        }
        return (String) writeColumnValueConverter.convert(entity, fieldValue);
    }


    public <ENTITY, SOURCE> StringColumnDefinition valueToString(Class<SOURCE> clazz, StringWriteColumnValueConverter<ENTITY, SOURCE> writeColumnValueConverter) {
        this.writeColumnValueConverter = writeColumnValueConverter;
        return this;
    }

    public <ENTITY, TARGET> StringColumnDefinition stringToValue(Class<TARGET> clazz, StringReadColumnValueConverter<ENTITY, TARGET> readColumnValueConverter) {
        this.readColumnValueConverter = readColumnValueConverter;
        return this;
    }

    public interface StringWriteColumnValueConverter<ENTITY, SOURCE> extends ColumnValueConverter<ENTITY, SOURCE, String> {

    }

    public interface StringReadColumnValueConverter<ENTITY, TARGET> extends ColumnValueConverter<ENTITY, String, TARGET> {

    }
}
