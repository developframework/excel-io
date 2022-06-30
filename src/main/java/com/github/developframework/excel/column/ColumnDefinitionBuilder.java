package com.github.developframework.excel.column;

import com.github.developframework.excel.ColumnDefinition;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 列定义构建器
 *
 * @author qiushui on 2019-05-18.
 */
public class ColumnDefinitionBuilder {

    private final Workbook workbook;

    private FormulaEvaluator formulaEvaluator;

    public ColumnDefinitionBuilder(Workbook workbook) {
        this.workbook = workbook;
    }

    @SafeVarargs
    public final <ENTITY> ColumnDefinition<ENTITY>[] columnDefinitions(ColumnDefinition<ENTITY>... columnDefinitions) {
        return columnDefinitions;
    }

    /**
     * 空列
     *
     * @param header 列名
     * @return 空列定义
     */
    public <ENTITY> BlankColumnDefinition<ENTITY> blank(String header) {
        return new BlankColumnDefinition<>(header);
    }

    public <ENTITY> BlankColumnDefinition<ENTITY> blank() {
        return blank(null);
    }

    /**
     * 通用列
     *
     * @param field  字段
     * @param header 列名
     * @return 字符串定义
     */
    public <ENTITY, FIELD> GeneralColumnDefinition<ENTITY, FIELD> column(String field, String header) {
        return new GeneralColumnDefinition<>(field, header);
    }

    public <ENTITY, FIELD> GeneralColumnDefinition<ENTITY, FIELD> column(String field) {
        return column(field, null);
    }

    /**
     * 公式列
     *
     * @param formula 公式字符串
     * @param header  列名
     * @return 公式列定义
     */
    public <ENTITY, FIELD> FormulaColumnDefinition<ENTITY, FIELD> formula(String field, String header, String formula) {
        if (formulaEvaluator == null) {
            formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
        }
        return new FormulaColumnDefinition<>(formulaEvaluator, field, header, formula);
    }

    public <ENTITY, FIELD> FormulaColumnDefinition<ENTITY, FIELD> formula(String header, String formula) {
        return formula(null, header, formula);
    }

    public <ENTITY, FIELD> FormulaColumnDefinition<ENTITY, FIELD> formula(String field) {
        return formula(field, null, null);
    }
}
