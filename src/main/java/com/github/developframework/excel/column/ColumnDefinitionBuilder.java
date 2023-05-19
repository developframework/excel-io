package com.github.developframework.excel.column;

import com.github.developframework.excel.ColumnDefinition;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.function.Function;

/**
 * 列定义构建器
 *
 * @author qiushui on 2019-05-18.
 */
public class ColumnDefinitionBuilder<ENTITY> {

    private final Workbook workbook;

    private FormulaEvaluator formulaEvaluator;

    public ColumnDefinitionBuilder(Workbook workbook) {
        this.workbook = workbook;
    }

    @SafeVarargs
    public final ColumnDefinition<ENTITY>[] columnDefinitions(ColumnDefinition<ENTITY>... columnDefinitions) {
        return columnDefinitions;
    }

    /**
     * 空列
     *
     * @param header 列名
     * @return 空列定义
     */
    public BlankColumnDefinition<ENTITY> blank(String header) {
        return new BlankColumnDefinition<>(header);
    }

    public BlankColumnDefinition<ENTITY> blank() {
        return blank(null);
    }

    /**
     * 通用列
     *
     * @param field  字段
     * @param header 列名
     * @return 字符串定义
     */
    public <FIELD> GeneralColumnDefinition<ENTITY, FIELD> column(String field, String header) {
        return new GeneralColumnDefinition<>(field, header);
    }

    public <FIELD> GeneralColumnDefinition<ENTITY, FIELD> column(String field) {
        return column(field, null);
    }

    /**
     * 字面量列
     *
     * @param field  字段
     * @param header 列名
     * @return 字符串定义
     */
    public LiteralColumnDefinition<ENTITY> literal(String field, String header) {
        return new LiteralColumnDefinition<>(field, header);
    }

    public LiteralColumnDefinition<ENTITY> literal(String field) {
        return literal(field, null);
    }

    /**
     * 公式列
     *
     * @param formula 公式字符串
     * @param header  列名
     * @return 公式列定义
     */
    public <FIELD> FormulaColumnDefinition<ENTITY, FIELD> formula(Class<?> fieldClass, String field, String header, String formula) {
        if (formulaEvaluator == null) {
            formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
        }
        return new FormulaColumnDefinition<>(formulaEvaluator, field, header, formula, null, fieldClass);
    }

    public <FIELD> FormulaColumnDefinition<ENTITY, FIELD> formula(Class<?> fieldClass, String header, String formula) {
        return formula(fieldClass, null, header, formula);
    }

    public <FIELD> FormulaColumnDefinition<ENTITY, FIELD> formula(Class<?> fieldClass, String field, String header, Function<Cell, String> formulaFunction) {
        if (formulaEvaluator == null) {
            formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
        }
        return new FormulaColumnDefinition<>(formulaEvaluator, field, header, null, formulaFunction, fieldClass);
    }

    public <FIELD> FormulaColumnDefinition<ENTITY, FIELD> formula(Class<?> fieldClass, String header, Function<Cell, String> formulaFunction) {
        return formula(fieldClass, null, header, formulaFunction);
    }

    public <FIELD> FormulaColumnDefinition<ENTITY, FIELD> formula(Class<?> fieldClass, String field) {
        return formula(fieldClass, field, null, (String) null);
    }
}
