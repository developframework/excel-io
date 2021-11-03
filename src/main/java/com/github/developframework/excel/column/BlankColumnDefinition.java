package com.github.developframework.excel.column;

import com.github.developframework.excel.ColumnDefinition;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

/**
 * @author qiushui on 2019-09-02.
 */
public class BlankColumnDefinition extends ColumnDefinition<String, Void> {

    public BlankColumnDefinition(String header) {
        super(null, null, header);
    }

    @Override
    protected CellType getColumnCellType() {
        return CellType.BLANK;
    }

    @Override
    protected void setCellValue(Cell cell, String convertValue) {
        cell.setCellValue(convertValue);
    }

    @Override
    protected String getCellValue(Cell cell) {
        return dataFormatter.formatCellValue(cell);
    }

    @Override
    protected String writeConvertValue(Object entity, Void fieldValue) {
        return null;
    }
}
