package com.github.developframework.excel;

import com.github.developframework.excel.column.ColumnDefinitionBuilder;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 表定义
 *
 * @author qiushui on 2019-05-18.
 */
public interface TableDefinition {

    /**
     * 配置表数据预处理器
     *
     * @return
     */
    default PreparedTableDataHandler<?, ?> preparedTableDataHandler() {
        return null;
    }

    /**
     * 是否有标题
     *
     * @return
     */
    default boolean hasTitle() {
        return false;
    }

    /**
     * 标题
     *
     * @return
     */
    default String title() {
        return null;
    }

    /**
     * 是否有列头
     */
    default boolean hasColumnHeader() {
        return true;
    }

    /**
     * 工作表名称
     */
    default String sheetName() {
        return null;
    }

    /**
     * 工作表
     */
    default Integer sheet() {
        return null;
    }

    /**
     * 表格位置
     */
    default TableLocation tableLocation() {
        return TableLocation.of(0, 0);
    }

    /**
     * 列定义
     */
    ColumnDefinition<?>[] columnDefinitions(Workbook workbook, ColumnDefinitionBuilder builder);

    default SheetExtraHandler<?> sheetExtraHandler() {
        return null;
    }
}
