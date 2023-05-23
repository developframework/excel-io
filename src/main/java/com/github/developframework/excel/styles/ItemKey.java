package com.github.developframework.excel.styles;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * @author qiushui on 2023-05-22.
 */
@FunctionalInterface
public interface ItemKey {

    void configureCellStyle(Workbook workbook, CellStyle cellStyle);
}
