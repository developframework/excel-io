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
     *
     * @param workbook
     * @return
     */
    public static CellStyle normalCellStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        borderThinAndVHCenter(cellStyle);
        return cellStyle;
    }

    /**
     * 生成数字型单元格风格
     *
     * @param workbook
     * @return
     */
    public static CellStyle numberCellStyle(Workbook workbook) {
        CellStyle numberCellStyle = workbook.createCellStyle();
        borderThinAndVHCenter(numberCellStyle);
        numberCellStyle.setDataFormat(workbook.createDataFormat().getFormat("0.00"));
        return numberCellStyle;
    }

    /**
     * 细边框并文本垂直水平居中
     *
     * @param cellStyle
     */
    private static void borderThinAndVHCenter(CellStyle cellStyle) {
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
    }
}
