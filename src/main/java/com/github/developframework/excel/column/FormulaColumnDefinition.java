package com.github.developframework.excel.column;

import com.github.developframework.excel.ColumnDefinition;
import develop.toolkit.base.utils.K;
import lombok.Getter;
import org.apache.poi.ss.usermodel.*;

/**
 * @author qiushui on 2019-09-02.
 */
public class FormulaColumnDefinition<FIELD> extends ColumnDefinition<Object, FIELD> {

    @Getter
    private String formula;

    private final FormulaEvaluator formulaEvaluator;

    public FormulaColumnDefinition(Workbook workbook, String field, String header) {
        super(workbook, field, header);
        this.formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
    }

    @Override
    protected CellType getColumnCellType() {
        return CellType.FORMULA;
    }

    @Override
    protected void setCellValue(Cell cell, Object convertValue) {
        cell.setCellFormula((String) convertValue);
    }

    @Override
    protected String writeConvertValue(Object entity, Object fieldValue) {
        return (String) fieldValue;
    }

    @Override
    protected Object getCellValue(Cell cell) {
        CellValue cellValue = formulaEvaluator.evaluate(cell);
        switch (cellValue.getCellType()) {
            case NUMERIC:
                return cellValue.getNumberValue();
            case STRING: {
                return K.map(cell.getStringCellValue(), String::trim);
            }
            case BOOLEAN:
                return cellValue.getBooleanValue();
            default:
                return null;
        }
    }

    public FormulaColumnDefinition<FIELD> formula(String formula) {
        this.formula = formula;
        return this;
    }
}
