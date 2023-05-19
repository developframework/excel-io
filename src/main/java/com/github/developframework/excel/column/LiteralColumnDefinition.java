package com.github.developframework.excel.column;

import com.github.developframework.excel.AbstractColumnDefinition;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * @author qiushui on 2023-01-31.
 */
public class LiteralColumnDefinition<ENTITY> extends AbstractColumnDefinition<ENTITY, Void> {

    private final String literal;

    public LiteralColumnDefinition(String literal, String header) {
        super(null, header);
        this.literal = literal;
    }

    @Override
    public Object writeIntoCell(Workbook workbook, Sheet sheet, Cell cell, ENTITY entity, int index) {
        final Object v;
        if (literal.equals("{no}")) {
            v = index + 1;
        } else {
            v = literal;
        }
        setCellValue(cell, v);
        return literal;
    }

    @Override
    public void readOutCell(Workbook workbook, Cell cell, ENTITY entity) {
        throw new IllegalStateException("不支持");
    }
}
