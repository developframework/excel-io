package com.github.developframework.excel.column;

import org.apache.poi.ss.usermodel.Workbook;

/**
 * 列定义构建器
 *
 * @author qiushui on 2019-01-16.
 */
public final class ColumnDefinitions {

    private Workbook workbook;

    public ColumnDefinitions(Workbook workbook) {
        this.workbook = workbook;
    }

    public BasicColumnDefinition basic(String fieldName) {
        return new BasicColumnDefinition(workbook, fieldName);
    }

    public DateTimeColumnDefinition dateTime(String fieldName) {
        return new DateTimeColumnDefinition(workbook, fieldName);
    }

    public NumberColumnDefinition number(String fieldName) {
        return new NumberColumnDefinition(workbook, fieldName);
    }

    public FormulaColumnDefinition formula(String fieldName) {
        return new FormulaColumnDefinition(workbook, fieldName);
    }

    public MultipleValueColumnDefinition multipleValue(String fieldName) {
        return new MultipleValueColumnDefinition(workbook, fieldName);
    }
}
