package com.github.developframework.excel;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

public interface CellStyleProvider {

    CellStyle provide(Workbook workbook, CellStyle originalStyle, Object value);
}
