package com.github.developframework.excel.styles;

import org.apache.poi.ss.usermodel.*;

/**
 * 默认的单元格风格
 *
 * @author qiushui on 2019-01-14.
 */
public final class DefaultCellStyles {

    /**
     * 普通单元格风格
     */
    public static CellStyle normalCellStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        borderThinAndVHCenter(cellStyle);
        return cellStyle;
    }

    /**
     * 生成数字型单元格风格
     */
    public static CellStyle numberCellStyle(Workbook workbook) {
        final CellStyle cellStyle = normalCellStyle(workbook);
        cellStyle.setDataFormat(workbook.createDataFormat().getFormat("0.00"));
        return cellStyle;
    }

    /**
     * 细边框并文本垂直水平居中
     */
    public static void borderThinAndVHCenter(CellStyle cellStyle) {
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
    }

    /**
     * 默认日期格式
     */
    public static CellStyle normalDateTimeCellStyle(Workbook workbook) {
        final CellStyle cellStyle = normalCellStyle(workbook);
        cellStyle.setDataFormat(workbook.createDataFormat().getFormat("yyyy-mm-dd hh:mm:ss") /* Excel的占位符格式，并非写错 */);
        return cellStyle;
    }

    /**
     * 设置对齐方式
     */
    public static void alignment(CellStyle cellStyle, HorizontalAlignment horizontalAlignment, VerticalAlignment verticalAlignment) {
        if (horizontalAlignment != null) {
            cellStyle.setAlignment(horizontalAlignment);
        }
        if (verticalAlignment != null) {
            cellStyle.setVerticalAlignment(verticalAlignment);
        }
    }
}
