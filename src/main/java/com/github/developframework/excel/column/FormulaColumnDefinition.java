package com.github.developframework.excel.column;

import com.github.developframework.excel.AbstractColumnDefinition;
import com.github.developframework.excel.ValueConvertUtils;
import lombok.Getter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;

/**
 * @author qiushui on 2019-09-02.
 */
public class FormulaColumnDefinition<ENTITY, FIELD> extends AbstractColumnDefinition<ENTITY, FIELD> {

    @Getter
    private final String formula;

    private final FormulaEvaluator formulaEvaluator;

    public FormulaColumnDefinition(FormulaEvaluator formulaEvaluator, String field, String header, String formula) {
        super(field, header);
        this.formula = formula;
        this.formulaEvaluator = formulaEvaluator;
    }

    @Override
    protected void setCellValue(Cell cell, Object convertValue) {
        cell.setCellFormula(
                formula
                        .replaceAll("\\{\\s*row\\s*}", String.valueOf(cell.getRowIndex() + 1))
                        .replaceAll("\\{\\s*column\\s*}", String.valueOf(cell.getColumnIndex() + 1))
        );
    }

    @Override
    protected Object getCellValue(Cell cell, Class<?> fieldClass) {
        final CellType cellType = formulaEvaluator.evaluateFormulaCell(cell);
        final CellValue cellValue = formulaEvaluator.evaluate(cell);
        switch (cellType) {
            case NUMERIC:
                return ValueConvertUtils.doubleConvert(cellValue.getNumberValue(), fieldClass);
            case BOOLEAN:
                return ValueConvertUtils.booleanConvert(cellValue.getBooleanValue(), fieldClass);
            case STRING:
                return cellValue.getStringValue();
            default:
                return null;
        }
    }
}
