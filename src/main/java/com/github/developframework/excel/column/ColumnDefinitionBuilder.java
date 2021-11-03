package com.github.developframework.excel.column;

import com.github.developframework.excel.ColumnDefinition;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 列定义构建器
 *
 * @author qiushui on 2019-05-18.
 */
public class ColumnDefinitionBuilder {

    private final Workbook workbook;

    public ColumnDefinitionBuilder(Workbook workbook) {
        this.workbook = workbook;
    }

    public ColumnDefinition<?, ?>[] columnDefinitions(ColumnDefinition<?, ?>... columnDefinitions) {
        return columnDefinitions;
    }

    /**
     * 空列
     *
     * @param header 列名
     * @return 空列定义
     */
    public BlankColumnDefinition blank(String header) {
        return new BlankColumnDefinition(header);
    }

    public BlankColumnDefinition blank() {
        return new BlankColumnDefinition(null);
    }

    /**
     * 字符串列
     *
     * @param field  字段
     * @param header 列名
     * @return 字符串定义
     */
    public <FIELD> StringColumnDefinition<FIELD> string(String field, String header) {
        return new StringColumnDefinition<>(workbook, field, header);
    }

    public <FIELD> StringColumnDefinition<FIELD> string(String field) {
        return new StringColumnDefinition<>(workbook, field, null);
    }

    /**
     * 数值列
     *
     * @param field  字段
     * @param header 列名
     * @return 数值列定义
     */
    public <FIELD> NumericColumnDefinition<FIELD> numeric(String field, String header) {
        return new NumericColumnDefinition<>(workbook, field, header);
    }

    public <FIELD> NumericColumnDefinition<FIELD> numeric(String field) {
        return new NumericColumnDefinition<>(workbook, field, null);
    }

    /**
     * 公式列
     *
     * @param formula 公式字符串
     * @param header  列名
     * @return 公式列定义
     */
    public <FIELD> FormulaColumnDefinition<FIELD> formula(String formula, String header) {
        return new FormulaColumnDefinition<>(workbook, formula, header);
    }

    public <FIELD> FormulaColumnDefinition<FIELD> formula(String formula) {
        return new FormulaColumnDefinition<>(workbook, formula, null);
    }
}
