package com.github.developframework.excel;

import com.github.developframework.excel.column.ColumnDefinition;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 表格定义
 *
 * @author qiushui on 2018-10-09.
 * @since 0.1
 */
public interface TableDefinition {

    /**
     * 是否有表头
     */
    boolean hasHeader();

    /**
     * 工作表名称
     */
    String sheetName();

    /**
     * 工作表
     */
    Integer sheet();

    /**
     * 起始列
     */
    int column();

    /**
     * 起始行
     */
    int row();

    /**
     * 列定义
     */
    ColumnDefinition[] columnDefinitions(Workbook workbook);

    /**
     * 设置表头单元格格式
     */
    void tableHeaderCellStyle(Workbook workbook, CellStyle cellStyle);

}
