package com.github.developframework.excel.column;

import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 基本列定义
 *
 * @author qiushui on 2018-10-10.
 * @since 0.1
 */
public class BasicColumnDefinition extends ColumnDefinition {

    public BasicColumnDefinition(Workbook workbook, String fieldName) {
        super(workbook);
        this.fieldName = fieldName;
        this.cellType = CellType.STRING;
    }

    @Override
    public void dealFillData(Cell cell, Object value) {
        cell.setCellValue(value.toString());
    }

    @Override
    public void dealReadData(Cell cell, Object instance) {
        String value = cell.getStringCellValue();
        Object object = readColumnValueConverter.map(converter -> converter.convert(instance, value)).orElse(value);
        try {
            FieldUtils.writeDeclaredField(instance, fieldName, object, true);
        } catch (IllegalAccessException e) {
            throw new RuntimeException(e);
        }
    }
}
