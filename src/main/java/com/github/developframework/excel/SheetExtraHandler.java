package com.github.developframework.excel;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

/**
 * 工作表扩展处理
 *
 * @param <T>
 */
public interface SheetExtraHandler<T> {

    void handle(Workbook workbook, Sheet sheet, int firstRow, int lastRow, List<T> list);
}
