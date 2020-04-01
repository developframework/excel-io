package com.github.developframework.excel;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

public interface SheetExtraHandler<T> {

    void handle(Workbook workbook, Sheet sheet, int firstRow, int lastRow, List<T> list);
}
