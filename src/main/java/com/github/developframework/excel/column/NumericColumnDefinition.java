package com.github.developframework.excel.column;

import com.github.developframework.excel.ColumnDefinition;
import com.github.developframework.excel.ColumnValueConverter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Workbook;

import java.math.BigDecimal;

/**
 * @author qiushui on 2019-05-19.
 */
public class NumericColumnDefinition extends ColumnDefinition<Double> {

    private NumericWriteColumnValueConverter writeColumnValueConverter;

    private NumericReadColumnValueConverter readColumnValueConverter;

    public NumericColumnDefinition(Workbook workbook, String field, String header) {
        super(workbook, field, header);
    }

    @Override
    protected CellType getColumnCellType() {
        return CellType.NUMERIC;
    }

    @Override
    protected void setCellValue(Cell cell, Double convertValue) {
        cell.setCellValue(convertValue);
    }

    @Override
    protected Double getCellValue(Cell cell) {
        return cell.getNumericCellValue();
    }

    @SuppressWarnings("unchecked")
    @Override
    protected Double writeConvertValue(Object entity, Object fieldValue) {
        if (writeColumnValueConverter != null) {
            fieldValue = writeColumnValueConverter.convert(entity, fieldValue);
        }
        if (fieldValue == null) {
            return null;
        } else if (fieldValue instanceof Integer) {
            return ((Integer) fieldValue).doubleValue();
        } else if (fieldValue instanceof Long) {
            return ((Long) fieldValue).doubleValue();
        } else if (fieldValue instanceof BigDecimal) {
            return ((BigDecimal) fieldValue).doubleValue();
        } else if (fieldValue instanceof Float) {
            return ((Float) fieldValue).doubleValue();
        } else if (fieldValue instanceof Double) {
            return (Double) fieldValue;
        } else {
            throw new IllegalArgumentException("must be Number Instance");
        }
    }

    @SuppressWarnings("unchecked")
    @Override
    protected <T> Object readConvertValue(Object entity, Double cellValue, Class<T> fieldClass) {
        Number convertValue;
        if (readColumnValueConverter != null) {
            convertValue = (Number) readColumnValueConverter.convert(entity, cellValue);
        } else {
            convertValue = cellValue;
        }
        if (convertValue == null) {
            return null;
        } else if (fieldClass == convertValue.getClass()) {
            return convertValue;
        } else if (fieldClass == String.class) {
            return convertValue.toString();
        } else if (fieldClass == Integer.class || fieldClass == int.class) {
            return convertValue.intValue();
        } else if (fieldClass == Long.class || fieldClass == long.class) {
            return convertValue.longValue();
        } else if (fieldClass == BigDecimal.class) {
            return new BigDecimal(convertValue.toString());
        } else if (fieldClass == Float.class || fieldClass == float.class) {
            return convertValue.floatValue();
        } else if (fieldClass == Double.class || fieldClass == double.class) {
            return convertValue.doubleValue();
        } else {
            throw new IllegalArgumentException("can not convert from \"java.lang.Double\" to \"" + fieldClass.getName() + "\"");
        }
    }

    public <ENTITY, SOURCE> NumericColumnDefinition valueToDouble(Class<SOURCE> clazz, NumericWriteColumnValueConverter<ENTITY, SOURCE> writeColumnValueConverter) {
        this.writeColumnValueConverter = writeColumnValueConverter;
        return this;
    }

    public <ENTITY, TARGET extends Number> NumericColumnDefinition doubleToValue(Class<TARGET> clazz, NumericReadColumnValueConverter<ENTITY, TARGET> readColumnValueConverter) {
        this.readColumnValueConverter = readColumnValueConverter;
        return this;
    }

    public interface NumericWriteColumnValueConverter<ENTITY, SOURCE> extends ColumnValueConverter<ENTITY, SOURCE, Double> {

    }

    public interface NumericReadColumnValueConverter<ENTITY, TARGET extends Number> extends ColumnValueConverter<ENTITY, Double, TARGET> {

    }
}
