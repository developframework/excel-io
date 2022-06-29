package com.github.developframework.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * @author qiushui on 2022-06-28.
 */
public interface ColumnDefinition<ENTITY> {

    /**
     * 列信息
     */
    default ColumnInfo getColumnInfo() {
        return null;
    }

    /**
     * 写入单元格
     *
     * @return 字段值
     */
    default Object writeIntoCell(Workbook workbook, Cell cell, ENTITY entity) {
        return null;
    }

    /**
     * 从单元格读取
     */
    default void readOutCell(Workbook workbook, Cell cell, ENTITY entity) {
    }

    /**
     * 配置单元格格式
     */
    default void configureCellStyle(Cell cell, CellStyleManager cellStyleManager, Object value) {
    }
}
