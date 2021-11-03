package com.github.developframework.excel.column;

import com.github.developframework.excel.ColumnDefinition;
import develop.toolkit.base.utils.K;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.function.BiFunction;

/**
 * 字符串列
 *
 * @author qiushui on 2019-05-18.
 */
public class StringColumnDefinition<FIELD> extends ColumnDefinition<String, FIELD> {

    protected StringColumnDefinition(Workbook workbook, String field, String header) {
        super(workbook, field, header);
    }

    @Override
    protected CellType getColumnCellType() {
        return CellType.STRING;
    }

    @Override
    protected void setCellValue(Cell cell, String convertValue) {
        cell.setCellValue(convertValue);
    }

    @Override
    protected String getCellValue(Cell cell) {
        return K.map(dataFormatter.formatCellValue(cell), String::trim);
    }

    @Override
    protected String writeConvertValue(Object entity, FIELD fieldValue) {
        if (writeConvertFunction == null) {
            return K.map(fieldValue, Object::toString);
        }
        return writeConvertFunction.apply(entity, fieldValue);
    }

    @SuppressWarnings("unused")
    public StringColumnDefinition<FIELD> valueToString(BiFunction<Object, FIELD, String> function) {
        this.writeConvertFunction = function;
        return this;
    }

    @SuppressWarnings("unused")
    public StringColumnDefinition<FIELD> stringToValue(BiFunction<Object, String, FIELD> function) {
        this.readConvertFunction = function;
        return this;
    }
}
