package com.github.developframework.excel.styles;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 默认的单元格样式
 *
 * @author qiushui on 2019-01-14.
 */
public final class DefaultCellStyles {

    // 标题单元格样式
    public static final String STYLE_TITLE = "font {size: 16; bold}";
    // 列头单元格样式
    public static final String STYLE_HEADER = "font {bold}";
    // 正文单元格样式
    public static final String STYLE_BODY = "";
    // 正文单元格样式 加粗
    public static final String STYLE_BODY_BOLD = "font {bold}";
    // 正文单元格样式 斜体
    public static final String STYLE_BODY_ITALIC = "font {italic}";
    // 正文单元格样式 2位百分比
    public static final String STYLE_BODY_PERCENT = "dataFormat {format: '0.00%'}";
    // 正文单元格样式 日期时间 （Excel的占位符格式，并非写错）
    public static final String STYLE_BODY_DATETIME = "dataFormat {format: 'yyyy-mm-dd hh:mm:ss'}";
    // 正文单元格样式 日期
    public static final String STYLE_BODY_DATE = "dataFormat {format: 'yyyy-mm-dd'}";
    // 正文单元格样式 时间 （Excel的占位符格式，并非写错）
    public static final String STYLE_BODY_TIME = "dataFormat {format: 'hh:mm:ss'}";
    // 正文单元格样式 数值右对齐
    public static final String STYLE_BODY_NUMBER = "align {horizontal: RIGHT}";
    // 正文单元格样式 数值右对齐 精度2
    public static final String STYLE_BODY_NUMBER_2_PRECISION = "align {horizontal: RIGHT} dataFormat {format: '0.00'}";

    /**
     * 根据key构建
     */
    public static CellStyle buildByCellStyleKey(Workbook workbook, String key) {
        if(CellStyleKey.isCellStyleKey(key)) {
            final CellStyleKey cellStyleKey = CellStyleKey.parse(key);
            CellStyle cellStyle = workbook.createCellStyle();
            cellStyleKey.configureCellStyle(workbook, cellStyle);
            return cellStyle;
        }
        return null;
    }
}
