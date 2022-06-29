package com.github.developframework.excel.column;

import com.github.developframework.excel.AbstractColumnDefinition;
import lombok.Getter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;

/**
 * @author qiushui on 2019-09-02.
 */
public class FormulaColumnDefinition<ENTITY, FIELD> extends AbstractColumnDefinition<ENTITY, FIELD> {

    @Getter
    private String formula;

    private final FormulaEvaluator formulaEvaluator;

    public FormulaColumnDefinition(FormulaEvaluator formulaEvaluator, String field, String header) {
        super(field, header);
        this.formulaEvaluator = formulaEvaluator;
    }

    @Override
    protected Object getCellValue(Cell cell) {
        CellValue cellValue = formulaEvaluator.evaluate(cell);
        switch (cellValue.getCellType()) {
            case NUMERIC:
                return cellValue.getNumberValue();
            case STRING: {
                final String value = cell.getStringCellValue();
                return value != null ? value.trim() : null;
            }
            case BOOLEAN:
                return cellValue.getBooleanValue();
            default:
                return null;
        }
    }

    //    fieldValue = columnInfo.field
//            .replaceAll("\\{\\s*row\\s*}", String.valueOf(cell.getRowIndex() + 1))
//            .replaceAll("\\{\\s*column\\s*}", String.valueOf(cell.getColumnIndex() + 1));
    public FormulaColumnDefinition<ENTITY, FIELD> formula(String formula) {
        this.formula = formula;
        return this;
    }
}
