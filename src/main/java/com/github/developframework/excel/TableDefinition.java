package com.github.developframework.excel;

import com.github.developframework.excel.column.ColumnDefinitionBuilder;
import com.github.developframework.excel.styles.CellStyleManager;
import com.github.developframework.excel.styles.DefaultCellStyles;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.Collections;
import java.util.List;
import java.util.Map;
import java.util.function.BiConsumer;

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
     * 数据预处理器
     */
    default PreparedTableDataHandler<?> preparedTableDataHandler() {
        return null;
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
    default BiConsumer<Workbook, CellStyle> globalCellStylesHandle() {
        return null;
    }

    /**
     * 工作表扩展处理
     */
    default SheetExtraHandler<?> sheetExtraHandler() {
        return null;
    }

    /**
     * 创建表标题
     *
     * @param sheet            工作表
     * @param cellStyleManager 单元格样式管理器
     * @param rowIndex         行索引
     * @param startColumnIndex 开始列索引
     * @param title            标题
     * @param columnSize       列数量
     */
    default void createTableTitle(Sheet sheet, CellStyleManager cellStyleManager, int rowIndex, final int startColumnIndex, String title, int columnSize) {
        if (title != null && !title.isBlank()) {
            Row titleRow = sheet.createRow(rowIndex);
            final CellStyle cellStyle = cellStyleManager.getCellStyle(DefaultCellStyles.STYLE_NORMAL_TITLE);
            for (int i = startColumnIndex; i < startColumnIndex + columnSize; i++) {
                titleRow.createCell(i).setCellStyle(cellStyle);
            }
            if (columnSize > 1) {
                sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, startColumnIndex, startColumnIndex + columnSize - 1));
            }
            titleRow.getCell(startColumnIndex).setCellValue(title);
        }
    }

    /**
     * 创建列头
     *
     * @param sheet             工作表
     * @param cellStyleManager  单元格样式管理器
     * @param rowIndex          列索引
     * @param startColumnIndex  开始列索引
     * @param columnDefinitions 列定义数组
     * @return rowIndex 最终的行号
     */
    default int createTableColumnHeader(Sheet sheet, CellStyleManager cellStyleManager, int rowIndex, final int startColumnIndex, ColumnDefinition<ENTITY>[] columnDefinitions) {
        Row headerRow = sheet.createRow(rowIndex);
        final CellStyle headerCellStyle = cellStyleManager.getCellStyle(DefaultCellStyles.STYLE_NORMAL_BOLD_HEADER);
        ColumnDefinition<ENTITY> columnDefinition;
        for (int i = 0; i < columnDefinitions.length; i++) {
            final Cell headerCell = headerRow.createCell(startColumnIndex + i);
            columnDefinition = columnDefinitions[i];
            if (columnDefinition == null || columnDefinition.getColumnInfo() == null) {
                continue;
            }
            final ColumnInfo columnInfo = columnDefinition.getColumnInfo();
            headerCell.setCellStyle(headerCellStyle);
            headerCell.setCellValue(columnInfo.header);
        }
        return rowIndex + 1;
    }

    /**
     * 创建表内容
     *
     * @param sheet             工作表
     * @param rowIndex          行索引
     * @param startColumnIndex  开始列索引
     * @param columnDefinitions 列定义数组
     * @param list              实体列表
     */
    default void createTableBody(Workbook workbook, Sheet sheet, CellStyleManager cellStyleManager, int rowIndex, final int startColumnIndex, ColumnDefinition<ENTITY>[] columnDefinitions, List<ENTITY> list) {
        // 渲染单元格
        for (int i = 0; i < list.size(); i++) {
            ENTITY entity = list.get(i);
            Row row = sheet.createRow(rowIndex + i);
            for (int j = 0; j < columnDefinitions.length; j++) {
                final ColumnDefinition<ENTITY> columnDefinition = columnDefinitions[j];
                final Cell cell = row.createCell(startColumnIndex + j);
                // 设置字段值
                Object convertValue = columnDefinition.writeIntoCell(workbook, cell, entity, i);
//                if (columnDefinition instanceof FormulaColumnDefinition) {
//                    convertValue = ((FormulaColumnDefinition<?, ?>) columnDefinition).getCellValue(cell, null);
//                }
                // 设置单元格样式
                columnDefinition.configureCellStyle(cell, cellStyleManager, convertValue);
            }
        }
    }
}
