package com.github.developframework.excel;

import com.github.developframework.excel.column.ColumnDefinitionBuilder;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Collections;
import java.util.Map;
import java.util.function.Consumer;

/**
 * 表定义
 *
 * @author qiushui on 2019-05-18.
 */
public interface TableDefinition<ENTITY> {

    /**
     * 表格信息
     */
    default TableInfo tableInfo() {
        return new TableInfo();
    }

    /**
     * 列定义
     */
    ColumnDefinition<ENTITY>[] columnDefinitions(Workbook workbook, ColumnDefinitionBuilder builder);

    /**
     * 每个处理
     */
    default void each(ENTITY entity) {

    }

    /**
     * 自定义单元格样式
     */
    default Map<String, CellStyle> customCellStyles(Workbook workbook) {
        return Collections.emptyMap();
    }

    /**
     * 全局样式处理
     */
    default Consumer<CellStyle> globalCellStylesHandle() {
        return null;
    }

    default SheetExtraHandler<?> sheetExtraHandler() {
        return null;
    }
}
