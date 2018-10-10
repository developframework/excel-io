package com.github.developframework.excel;

import org.apache.poi.ss.usermodel.*;

/**
 * 最简表格定义
 *
 * @author qiushui on 2018-10-10.
 * @since 0.1
 */
public abstract class AbstractTableDefinition implements TableDefinition {

    @Override
    public boolean hasHeader() {
        return true;
    }

    @Override
    public String sheetName() {
        return null;
    }

    @Override
    public Integer sheet() {
        return null;
    }

    @Override
    public int column() {
        return 0;
    }

    @Override
    public int row() {
        return 0;
    }

    @Override
    public void tableHeaderCellStyle(Workbook workbook, CellStyle cellStyle) {
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        borderThin(cellStyle);
        Font font = workbook.createFont();
        font.setBold(true);
        cellStyle.setFont(font);
    }

    /**
     * 加边框
     *
     * @param cellStyle
     */
    protected void borderThin(CellStyle cellStyle) {
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
    }

}
