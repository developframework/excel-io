package com.github.developframework.excel.styles;

import org.apache.poi.ss.usermodel.*;

/**
 * 默认的单元格样式
 *
 * @author qiushui on 2019-01-14.
 */
public final class DefaultCellStyles {

    public static final String STYLE_NORMAL = "normal";

    public static final String STYLE_NORMAL_TITLE = "normalTitle";

    public static final String STYLE_NORMAL_DATETIME = "normalDateTime";
    public static final String STYLE_NORMAL_NUMBER = "normalNumber";

    public static final String STYLE_NORMAL_BOLD = "normalBold";

    public static final String STYLE_NORMAL_BOLD_HEADER = "normalBoldHeader";
    public static final String STYLE_NORMAL_PERCENT = "normalPercent";

    /**
     * 普通单元格风格
     */
    public static CellStyle normalCellStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        // 细边框并文本垂直水平居中文本
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setDataFormat(workbook.createDataFormat().getFormat("@"));
        return cellStyle;
    }

    /**
     * 生成数字型单元格风格 （右对齐）
     */
    public static CellStyle normalNumberCellStyle(Workbook workbook) {
        final CellStyle cellStyle = normalCellStyle(workbook);
        cellStyle.setAlignment(HorizontalAlignment.RIGHT);
        return cellStyle;
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
     * 百分比
     */
    public static CellStyle normalPercentCellStyle(Workbook workbook) {
        final CellStyle cellStyle = normalCellStyle(workbook);
        cellStyle.setDataFormat(workbook.createDataFormat().getFormat("0.00%") /* Excel的占位符格式，并非写错 */);
        return cellStyle;
    }

    /**
     * 默认加粗格式
     */
    public static CellStyle normalBoldCellStyle(Workbook workbook) {
        final CellStyle cellStyle = normalCellStyle(workbook);
        Font font = workbook.createFont();
        font.setBold(true);
        cellStyle.setFont(font);
        return cellStyle;
    }
}
