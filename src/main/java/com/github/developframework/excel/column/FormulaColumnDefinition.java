package com.github.developframework.excel.column;

import org.apache.poi.ss.usermodel.*;

/**
 * @author qiushui on 2018-10-25.
 */
public class FormulaColumnDefinition extends ColumnDefinition {

    private Workbook workbook;

    private String formula;

    public FormulaColumnDefinition(Workbook workbook, String formula) {
        this.workbook = workbook;
        this.formula = formula;
        this.cellStyle = workbook.createCellStyle();
        this.cellStyle.setAlignment(HorizontalAlignment.CENTER);
        this.cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        borderThin(cellStyle);
        cellType = CellType.FORMULA;
    }

    @Override
    public void dealFillData(Cell cell, Object row) {
        cell.setCellFormula(formula.replaceAll("\\{row\\}", String.valueOf(row)));
    }

    @Override
    public void dealReadData(Cell cell, Object instance) {
//        Object object = readColumnValueConverter.map(converter -> converter.convert(instance, value)).orElse(value);
    }

    /**
     * 设置格式
     *
     * @param format
     * @return
     */
    public FormulaColumnDefinition format(String format) {
        DataFormat dataFormat = workbook.createDataFormat();
        this.cellStyle.setDataFormat(dataFormat.getFormat(format));
        return this;
    }
}
