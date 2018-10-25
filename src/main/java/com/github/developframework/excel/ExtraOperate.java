package com.github.developframework.excel;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * @author qiushui on 2018-10-25.
 */
public interface ExtraOperate {

    void operate(Workbook workbook, Sheet sheet);
}
