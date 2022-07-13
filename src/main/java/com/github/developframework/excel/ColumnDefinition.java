package com.github.developframework.excel;

import com.github.developframework.excel.styles.CellStyleManager;
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
     * 描述如何把值写入单元格
     *
     * @return 字段值
     */
    default Object writeIntoCell(Workbook workbook, Cell cell, ENTITY entity) {
        return null;
    }

    /**
     * 描述如何从单元格读取值并装填到实体
     */
    default void readOutCell(Workbook workbook, Cell cell, ENTITY entity) {
    }

    /**
     * 配置单元格格式
     */
    default void configureCellStyle(Cell cell, CellStyleManager cellStyleManager, Object value) {
    }
}
