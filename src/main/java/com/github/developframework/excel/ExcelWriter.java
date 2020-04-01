package com.github.developframework.excel;

import com.github.developframework.excel.column.BlankColumnDefinition;
import com.github.developframework.excel.column.ColumnDefinitionBuilder;
import com.github.developframework.excel.column.FormulaColumnDefinition;
import com.github.developframework.excel.styles.DefaultCellStyles;
import com.github.developframework.expression.ExpressionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
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
    public <ENTITY> ExcelWriter load(List<ENTITY> data, TableDefinition tableDefinition) {
        writeInternal(tableDefinition, data);
        return this;
    }

    /**
     * 填充数据
     *
     * @param data 实体列表
     * @param tableDefinition 表定义
     * @return 写出器
     */
    public <ENTITY> ExcelWriter load(ENTITY[] data, TableDefinition tableDefinition) {
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
    public void writeToFile(String filename) {
        try (OutputStream os = new FileOutputStream(filename)) {
            write(os);
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
        try {
            ByteArrayOutputStream os = new ByteArrayOutputStream();
            write(os);
            byte[] bytes = os.toByteArray();
            os.close();
            return bytes;
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 写入表格
     *
     * @param tableDefinition 标定仪
     * @param list 实体列表
     * @param <ENTITY> 实体类泛型
     */
    @SuppressWarnings({"unchecked", "rawtypes"})
    private <ENTITY> void writeInternal(TableDefinition tableDefinition, List<ENTITY> list) {
        final Sheet sheet = createSheet(tableDefinition);
        final TableLocation tableLocation = tableDefinition.tableLocation();
        PreparedTableDataHandler<ENTITY, ?> preparedTableDataHandler = (PreparedTableDataHandler<ENTITY, ?>) tableDefinition.preparedTableDataHandler();
        final List<?> finalList = preparedTableDataHandler == null ? list : preparedTableDataHandler.handle(list);
        final ColumnDefinition<?>[] columnDefinitions = tableDefinition.columnDefinitions(workbook, new ColumnDefinitionBuilder(workbook));
        final int startColumnIndex = tableLocation.getColumn();

        int rowIndex = tableLocation.getRow();
        if (tableDefinition.hasTitle() && tableDefinition.title() != null) {
            createTableTitle(sheet, rowIndex++, startColumnIndex, tableDefinition.title(), columnDefinitions.length);
        }
        if (tableDefinition.hasColumnHeader()) {
            createTableColumnHeader(sheet, rowIndex++, startColumnIndex, columnDefinitions);
        }
        createTableBody(sheet, rowIndex, startColumnIndex, columnDefinitions, finalList);
        SheetExtraHandler sheetExtraHandler = tableDefinition.sheetExtraHandler();
        if (sheetExtraHandler != null) {
            sheetExtraHandler.handle(workbook, sheet, rowIndex, rowIndex + list.size(), list);
        }
    }

    /**
     * 创建工作表
     *
     * @param tableDefinition 表定义
     * @return 工作表
     */
    private Sheet createSheet(TableDefinition tableDefinition) {
        if (tableDefinition.sheetName() == null) {
            return workbook.createSheet();
        } else {
            return workbook.createSheet(tableDefinition.sheetName());
        }
    }

    /**
     * 创建表标题
     *
     * @param sheet 工作表
     * @param rowIndex 行索引
     * @param startColumnIndex 开始列索引
     * @param title 标题
     * @param columnSize 列数量
     */
    private void createTableTitle(Sheet sheet, int rowIndex, final int startColumnIndex, String title, int columnSize) {
        if (StringUtils.isNotEmpty(title)) {
            Row titleRow = sheet.createRow(rowIndex);
            for (int i = startColumnIndex; i < startColumnIndex + columnSize; i++) {
                titleRow.createCell(i).setCellStyle(DefaultCellStyles.normalCellStyle(workbook));
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
     * @param sheet 工作表
     * @param rowIndex 列索引
     * @param startColumnIndex 开始列索引
     * @param columnDefinitions 列定义数组
     */
    private void createTableColumnHeader(Sheet sheet, int rowIndex, final int startColumnIndex, ColumnDefinition<?>[] columnDefinitions) {
        Row headerRow = sheet.createRow(rowIndex);
        CellStyle headerCellStyle = DefaultCellStyles.normalCellStyle(workbook);
        for (int i = 0; i < columnDefinitions.length; i++) {
            Cell headerCell = headerRow.createCell(startColumnIndex + i);
            ColumnDefinition<?> columnDefinition = columnDefinitions[i];
            if (columnDefinition == null) {
                continue;
            }
            headerCell.setCellStyle(headerCellStyle);
            if (columnDefinition.getHeader() != null) {
                headerCell.setCellValue(columnDefinition.getHeader());
            }
        }
    }

    /**
     * 创建表内容
     *
     * @param sheet 工作表
     * @param rowIndex 行索引
     * @param startColumnIndex 开始列索引
     * @param columnDefinitions 列定义数组
     * @param list 实体列表
     * @param <ENTITY> 实体类泛型
     */
    private <ENTITY> void createTableBody(Sheet sheet, int rowIndex, final int startColumnIndex, ColumnDefinition<?>[] columnDefinitions, List<ENTITY> list) {
        // 默认单元格风格
        final CellStyle DEFAULT_CELL_STYLE = DefaultCellStyles.normalCellStyle(workbook);
        // 渲染单元格
        for (int i = 0; i < list.size(); i++) {
            ENTITY entity = list.get(i);
            Row row = sheet.createRow(rowIndex + i);
            for (int j = 0; j < columnDefinitions.length; j++) {
                Cell cell = row.createCell(startColumnIndex + j);
                ColumnDefinition<?> columnDefinition = columnDefinitions[j];
                if (columnDefinition == null || columnDefinition instanceof BlankColumnDefinition) {
                    cell.setCellStyle(configCellStyle(columnDefinitions[j], DEFAULT_CELL_STYLE, null));
                    continue;
                }
                Object fieldValue;
                if (columnDefinition instanceof FormulaColumnDefinition) {
                    fieldValue = columnDefinition
                            .getField()
                            .replaceAll("\\{\\s*row\\s*}", String.valueOf(cell.getRowIndex() + 1))
                            .replaceAll("\\{\\s*column\\s*}", String.valueOf(cell.getColumnIndex() + 1));
                } else {
                    fieldValue = ExpressionUtils.getValue(entity, columnDefinition.field);
                }
                cell.setCellStyle(configCellStyle(columnDefinitions[j], DEFAULT_CELL_STYLE, fieldValue));
                columnDefinition.writeIntoCell(entity, cell, fieldValue);
            }
        }
        // 设置列宽
        for (int i = 0; i < columnDefinitions.length; i++) {
            if (columnDefinitions[i].columnWidth != null) {
                sheet.setColumnWidth(startColumnIndex + i, columnDefinitions[i].columnWidth * 256);
            }
        }
    }

    /**
     * 设置单元格风格
     *
     * @param columnDefinition 列定义
     * @return 单元格风格
     */
    private CellStyle configCellStyle(ColumnDefinition<?> columnDefinition, CellStyle defaultCellStyle, Object value) {
        CellStyle cellStyle = columnDefinition.cellStyleProvider == null ? defaultCellStyle : columnDefinition.cellStyleProvider.provide(workbook, defaultCellStyle, value);
        if (columnDefinition.format != null) {
            if (cellStyle == defaultCellStyle) {
                cellStyle = DefaultCellStyles.normalCellStyle(workbook);
            }
            cellStyle.setDataFormat(workbook.createDataFormat().getFormat(columnDefinition.format));
        }
        if (columnDefinition.alignment != null) {
            if (cellStyle == defaultCellStyle) {
                cellStyle = DefaultCellStyles.normalCellStyle(workbook);
            }
            cellStyle.setAlignment(columnDefinition.alignment.getFirstValue());
            cellStyle.setVerticalAlignment(columnDefinition.alignment.getSecondValue());
        }
        return cellStyle;
    }
}
