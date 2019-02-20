package com.github.developframework.excel.column;

import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Workbook;

import java.lang.reflect.Field;
import java.math.BigDecimal;

/**
 * 数值列定义
 *
 * @author qiushui on 2018-10-10.
 * @since 0.1
 */
public class NumberColumnDefinition extends BasicColumnDefinition {

    public NumberColumnDefinition(Workbook workbook, String field) {
        super(workbook, field);
        this.cellType = CellType.NUMERIC;
    }

    public NumberColumnDefinition(Workbook workbook, String field, String format) {
        super(workbook, field);
        this.cellType = CellType.NUMERIC;
        if (StringUtils.isNotBlank(format)) {
            DataFormat dataFormat = workbook.createDataFormat();
            this.cellStyle.setDataFormat(dataFormat.getFormat(format));
        }
    }

    @Override
    public void dealFillData(Cell cell, Object value) {
        cell.setCellValue(Double.parseDouble(value.toString()));
    }

    @Override
    public void dealReadData(Cell cell, Object instance) {
        Class<?> instanceClass = instance.getClass();
        Field field = FieldUtils.getField(instanceClass, fieldName, true);
        Class<?>[] acceptClasses = new Class<?>[]{
             Integer.class, int.class,
             Long.class, long.class, BigDecimal.class
        };

        if (ArrayUtils.contains(acceptClasses, field.getType())) {
            try {
                double value = cell.getNumericCellValue();
                Object object = readColumnValueConverter.map(converter -> converter.convert(instance, value)).orElse(value);
                FieldUtils.writeDeclaredField(instance, fieldName, object, true);
            } catch (IllegalAccessException e) {
                throw new RuntimeException(e);
            }
        }
    }

    /**
     * 设置格式
     *
     * @param format
     * @return
     */
    public ColumnDefinition format(String format) {
        if (StringUtils.isNotBlank(format)) {
            DataFormat dataFormat = workbook.createDataFormat();
            this.cellStyle.setDataFormat(dataFormat.getFormat(format));
        }
        return this;
    }
}
