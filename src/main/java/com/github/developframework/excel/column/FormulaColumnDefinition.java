package com.github.developframework.excel.column;

import com.github.developframework.excel.ColumnDefinition;
import com.github.developframework.excel.ColumnValueConverter;
import lombok.Getter;
import org.apache.poi.ss.usermodel.*;

import java.math.BigDecimal;

/**
 * @author qiushui on 2019-09-02.
 */
public class FormulaColumnDefinition extends ColumnDefinition<Object> {

    @Getter
    private String formula;

    private FormulaEvaluator formulaEvaluator;

    private FormulaReadColumnValueConverter readColumnValueConverter;

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
            case STRING:
                return cellValue.getStringValue();
            case BOOLEAN:
                return cellValue.getBooleanValue();
            default:
                return null;
        }
    }

    @Override
    @SuppressWarnings("unchecked")
    protected <T> Object readConvertValue(Object entity, Object cellValue, Class<T> fieldClass) {
        Object convertValue;
        if (readColumnValueConverter != null) {
            convertValue = readColumnValueConverter.convert(entity, cellValue);
        } else {
            convertValue = cellValue;
        }
        if (convertValue == null) {
            return null;
        } else if (fieldClass == cellValue.getClass()) {
            return convertValue;
        } else if (fieldClass == String.class) {
            return convertValue.toString();
        } else if (fieldClass == Integer.class || fieldClass == int.class) {
            return Double.valueOf(convertValue.toString()).intValue();
        } else if (fieldClass == Long.class || fieldClass == long.class) {
            return Double.valueOf(convertValue.toString()).longValue();
        } else if (fieldClass == BigDecimal.class) {
            return new BigDecimal(convertValue.toString());
        } else if (fieldClass == Float.class || fieldClass == float.class) {
            return Double.valueOf(convertValue.toString()).floatValue();
        } else if (fieldClass == Double.class || fieldClass == double.class) {
            return Double.valueOf(convertValue.toString());
        } else {
            throw new IllegalArgumentException("can not convert from \"java.lang.String\" to \"" + fieldClass.getName() + "\"");
        }
    }

    public <ENTITY, TARGET> FormulaColumnDefinition stringToValue(Class<TARGET> clazz, FormulaReadColumnValueConverter<ENTITY, TARGET> readColumnValueConverter) {
        this.readColumnValueConverter = readColumnValueConverter;
        return this;
    }

    public FormulaColumnDefinition formula(String formula) {
        this.formula = formula;
        return this;
    }

    public interface FormulaReadColumnValueConverter<ENTITY, TARGET> extends ColumnValueConverter<ENTITY, Object, TARGET> {

    }
}
