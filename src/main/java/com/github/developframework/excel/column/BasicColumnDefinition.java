package com.github.developframework.excel.column;

import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.poi.ss.usermodel.*;

/**
 * 基本列定义
 *
 * @author qiushui on 2018-10-10.
 * @since 0.1
 */
public class BasicColumnDefinition extends ColumnDefinition {

    public BasicColumnDefinition(Workbook workbook, String header, String fieldName) {
        this.header = header;
        this.fieldName = fieldName;
        this.cellStyle = workbook.createCellStyle();
        this.cellStyle.setAlignment(HorizontalAlignment.CENTER);
        this.cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        borderThin(cellStyle);
        this.cellType = CellType.STRING;
    }

    @Override
    public void dealFillData(Cell cell, Object value) {
        cell.setCellValue(value.toString());
    }

    @Override
    public void dealReadData(Cell cell, Object instance) {
        String value = cell.getStringCellValue();
        try {
            FieldUtils.writeDeclaredField(instance, fieldName, value, true);
        } catch (IllegalAccessException e) {
            throw new RuntimeException(e);
        }
    }
}
