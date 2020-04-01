package com.github.developframework.excel;

import org.apache.poi.ss.usermodel.Workbook;

/**
 * @author qiushui on 2018-10-10.
 */
public abstract class ExcelProcessor {

    protected Workbook workbook;

    protected ExcelProcessor(Workbook workbook) {
        this.workbook = workbook;
    }
}
