package com.github.developframework.excel.column;

import com.github.developframework.excel.ColumnDefinition;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.function.BiFunction;

/**
 * @author qiushui on 2019-05-19.
 */
public class NumericColumnDefinition<FIELD> extends ColumnDefinition<Number, FIELD> {

    public NumericColumnDefinition(Workbook workbook, String field, String header) {
        super(workbook, field, header);
    }

    @Override
    protected CellType getColumnCellType() {
        return CellType.NUMERIC;
    }

    @Override
    protected void setCellValue(Cell cell, Number convertValue) {
        cell.setCellValue(convertValue.doubleValue());
    }

    @Override
    protected Number getCellValue(Cell cell) {
        return cell.getNumericCellValue();
    }

    @Override
    protected Number writeConvertValue(Object entity, FIELD fieldValue) {
        if (writeConvertFunction != null) {
            return writeConvertFunction.apply(entity, fieldValue);
        }
        return (Number) fieldValue;
    }

    @SuppressWarnings({"unchecked", "unused"})
    public <ENTITY> NumericColumnDefinition<FIELD> valueToNumber(BiFunction<ENTITY, FIELD, Number> function) {
        this.writeConvertFunction = (BiFunction<Object, FIELD, Number>) function;
        return this;
    }

    @SuppressWarnings({"unchecked", "unused"})
    public <ENTITY> NumericColumnDefinition<FIELD> numberToValue(BiFunction<ENTITY, Number, FIELD> function) {
        this.readConvertFunction = (BiFunction<Object, Number, FIELD>) function;
        return this;
    }
}
