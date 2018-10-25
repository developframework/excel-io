package com.github.developframework.excel.column;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;

/**
 * @author qiushui on 2018-10-25.
 */
public class FormulaColumnDefinition extends ColumnDefinition {

    private String formula;

    public FormulaColumnDefinition(Workbook workbook, String header, String formula, String format, int maxLength) {
        this.header = header;
        this.formula = formula;
        this.cellStyle = workbook.createCellStyle();
        this.cellStyle.setAlignment(HorizontalAlignment.CENTER);
        this.cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        borderThin(cellStyle);
        cellType = CellType.FORMULA;
        if (StringUtils.isNotBlank(format)) {
            DataFormat dataFormat = workbook.createDataFormat();
            this.cellStyle.setDataFormat(dataFormat.getFormat(format));
        }
        this.maxLength = maxLength;
    }

    @Override
    public void dealFillData(Cell cell, Object row) {
        cell.setCellFormula(formula.replaceAll("\\{row\\}", String.valueOf(row)));
    }

    @Override
    public void dealReadData(Cell cell, Object instance) {

    }
}
