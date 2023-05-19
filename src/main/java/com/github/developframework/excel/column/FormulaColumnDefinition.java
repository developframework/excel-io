package com.github.developframework.excel.column;

import com.github.developframework.excel.AbstractColumnDefinition;
import com.github.developframework.excel.utils.ValueConvertUtils;
import lombok.Getter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;

import java.util.function.Function;

/**
 * @author qiushui on 2019-09-02.
 */
public class FormulaColumnDefinition<ENTITY, FIELD> extends AbstractColumnDefinition<ENTITY, FIELD> {

    @Getter
    private final String formula;

    private final Function<Cell, String> formulaFunction;

    private final Class<?> fieldClass;

    private final FormulaEvaluator formulaEvaluator;

    public FormulaColumnDefinition(FormulaEvaluator formulaEvaluator, String field, String header, String formula, Function<Cell, String> formulaFunction, Class<?> fieldClass) {
        super(field, header);
        this.formula = formula;
        this.formulaFunction = formulaFunction;
        this.formulaEvaluator = formulaEvaluator;
        this.fieldClass = fieldClass;
    }

    @Override
    protected void setCellValue(Cell cell, Object convertValue) {
        final String finalFormula = formulaFunction != null ? formulaFunction.apply(cell) : formula;
        cell.setCellFormula(
                finalFormula
                        .replaceAll("\\{\\s*row\\s*}", String.valueOf(cell.getRowIndex() + 1))
                        .replaceAll("\\{\\s*column\\s*}", String.valueOf(cell.getColumnIndex() + 1))
        );
    }

    @Override
    public Object getCellValue(Cell cell, Class<?> fieldClass) {
        final CellValue cellValue = formulaEvaluator.evaluate(cell);
        switch (cellValue.getCellType()) {
            case NUMERIC: {
                cell.setCellType(CellType.STRING);
                return ValueConvertUtils.doubleConvert(cell.getStringCellValue(), this.fieldClass);
            }
            case BOOLEAN:
                return ValueConvertUtils.booleanConvert(cellValue.getBooleanValue(), this.fieldClass);
            case STRING:
                return ValueConvertUtils.stringConvert(cellValue.getStringValue(), this.fieldClass);
            default:
                return null;
        }
    }
}
