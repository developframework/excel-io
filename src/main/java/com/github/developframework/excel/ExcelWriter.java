package com.github.developframework.excel;

import com.github.developframework.excel.column.ColumnDefinitionBuilder;
import com.github.developframework.excel.column.FormulaColumnDefinition;
import com.github.developframework.excel.styles.CellStyleManager;
import com.github.developframework.excel.styles.DefaultCellStyles;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;

import java.io.*;
import java.util.Arrays;
import java.util.List;

/**
 * @author qiushui on 2019-05-18.
 */
@SuppressWarnings("unused")
public class ExcelWriter extends ExcelProcessor {

    protected ExcelWriter(Workbook workbook) {
        super(workbook);
    }

    /**
     * 填充数据
     *
     * @param data            实体列表
     * @param tableDefinition 表定义
     * @return 写出器
     */
    public <ENTITY> ExcelWriter load(List<ENTITY> data, TableDefinition<ENTITY> tableDefinition) {
        writeInternal(tableDefinition, data);
        return this;
    }

    /**
     * 填充数据
     *
     * @param data            实体列表
     * @param tableDefinition 表定义
     * @return 写出器
     */
    public <ENTITY> ExcelWriter load(ENTITY[] data, TableDefinition<ENTITY> tableDefinition) {
        writeInternal(tableDefinition, Arrays.asList(data));
        return this;
    }

    /**
     * 写出
     */
    public void write(OutputStream outputStream) {
        try {
            workbook.write(outputStream);
            workbook.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 写出到文件
     *
     * @param filename 文件名
     */
    public File writeToFile(String filename) {
        File file = new File(filename);
        try (OutputStream os = new FileOutputStream(file)) {
            write(os);
            return file;
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 写出到字节数组
     *
     * @return 字节数组
     */
    public byte[] writeToByteArray() {
        try (ByteArrayOutputStream os = new ByteArrayOutputStream()) {
            write(os);
            return os.toByteArray();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 写入表格
     *
     * @param tableDefinition 标定仪
     * @param list            实体列表
     * @param <ENTITY>        实体类泛型
     */
    @SuppressWarnings({"unchecked", "rawtypes"})
    private <ENTITY> void writeInternal(TableDefinition<ENTITY> tableDefinition, List<ENTITY> list) {
        final TableInfo tableInfo = tableDefinition.tableInfo();

        // 创建工作表
        final Sheet sheet = createSheet(tableInfo);

        // 列定义
        final ColumnDefinition<ENTITY>[] columnDefinitions = tableDefinition.columnDefinitions(workbook, new ColumnDefinitionBuilder(workbook));

        // 表位置
        final int startColumnIndex = tableInfo.tableLocation.getColumn();
        int rowIndex = tableInfo.tableLocation.getRow();

        // 单元格样式管理器
        final CellStyleManager cellStyleManager = new CellStyleManager(workbook, tableDefinition);

        // 创建表头标题
        if (tableInfo.hasTitle && tableInfo.title != null) {
            createTableTitle(sheet, cellStyleManager, rowIndex++, startColumnIndex, tableInfo.title, columnDefinitions.length);
        }
        // 创建列头
        if (tableInfo.hasColumnHeader) {
            createTableColumnHeader(sheet, cellStyleManager, rowIndex++, startColumnIndex, columnDefinitions);
        }

        // 创建表主体数据
        createTableBody(sheet, cellStyleManager, rowIndex, startColumnIndex, columnDefinitions, list);

        // 设置列宽
        for (int i = 0; i < columnDefinitions.length; i++) {
            int col = startColumnIndex + i;
            final ColumnInfo columnInfo = columnDefinitions[i].getColumnInfo();
            if (columnInfo != null) {
                if (columnInfo.columnWidth == null) {
                    // 自动列宽
                    sheet.autoSizeColumn(col);
                    // 解决自动设置列宽中文失效的问题
                    sheet.setColumnWidth(col, sheet.getColumnWidth(i) * 17 / 10);
                } else {
                    sheet.setColumnWidth(col, columnInfo.columnWidth * 256);
                }
            }
        }

        // 工作表额外处理
        SheetExtraHandler sheetExtraHandler = tableDefinition.sheetExtraHandler();
        if (sheetExtraHandler != null) {
            sheetExtraHandler.handle(workbook, sheet, cellStyleManager, rowIndex, rowIndex + list.size(), list);
        }
    }

    /**
     * 创建工作表
     *
     * @param tableInfo 表格信息
     * @return 工作表
     */
    private Sheet createSheet(TableInfo tableInfo) {
        Sheet sheet;
        if (tableInfo.sheetName == null) {
            sheet = workbook.createSheet();
        } else {
            sheet = workbook.createSheet(tableInfo.sheetName);
        }
        if (sheet instanceof SXSSFSheet) {
            ((SXSSFSheet) sheet).setRandomAccessWindowSize(-1);
            // 开启追踪列宽
            ((SXSSFSheet) sheet).trackAllColumnsForAutoSizing();
        }
        return sheet;
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
    private void createTableTitle(Sheet sheet, CellStyleManager cellStyleManager, int rowIndex, final int startColumnIndex, String title, int columnSize) {
        if (title != null && !title.isBlank()) {
            Row titleRow = sheet.createRow(rowIndex);
            final CellStyle cellStyle = cellStyleManager.getCellStyle(DefaultCellStyles.STYLE_NORMAL);
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
     */
    private <ENTITY> void createTableColumnHeader(Sheet sheet, CellStyleManager cellStyleManager, int rowIndex, final int startColumnIndex, ColumnDefinition<ENTITY>[] columnDefinitions) {
        Row headerRow = sheet.createRow(rowIndex);
        final CellStyle headerCellStyle = cellStyleManager.getCellStyle(DefaultCellStyles.STYLE_NORMAL_BOLD);
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
    }

    /**
     * 创建表内容
     *
     * @param sheet             工作表
     * @param rowIndex          行索引
     * @param startColumnIndex  开始列索引
     * @param columnDefinitions 列定义数组
     * @param list              实体列表
     * @param <ENTITY>          实体类泛型
     */
    private <ENTITY> void createTableBody(Sheet sheet, CellStyleManager cellStyleManager, int rowIndex, final int startColumnIndex, ColumnDefinition<ENTITY>[] columnDefinitions, List<ENTITY> list) {
        // 渲染单元格
        for (int i = 0; i < list.size(); i++) {
            ENTITY entity = list.get(i);
            Row row = sheet.createRow(rowIndex + i);
            for (int j = 0; j < columnDefinitions.length; j++) {
                final ColumnDefinition<ENTITY> columnDefinition = columnDefinitions[j];
                final Cell cell = row.createCell(startColumnIndex + j);
                // 设置字段值
                Object convertValue = columnDefinition.writeIntoCell(workbook, cell, entity);
                if (columnDefinition instanceof FormulaColumnDefinition) {
                    convertValue = ((FormulaColumnDefinition<?, ?>) columnDefinition).getCellValue(cell, null);
                }
                // 设置单元格样式
                columnDefinition.configureCellStyle(cell, cellStyleManager, convertValue);
            }
        }
    }
}
