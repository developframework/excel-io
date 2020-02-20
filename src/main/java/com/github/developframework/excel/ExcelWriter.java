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
import java.util.stream.Stream;

/**
 * @author qiushui on 2019-05-18.
 */
public class ExcelWriter extends ExcelProcessor {

    protected ExcelWriter(Workbook workbook) {
        super(workbook);
    }

    /**
     * 填充数据
     *
     * @param data
     * @param tableDefinition
     * @return
     */
    public <ENTITY> ExcelWriter load(List<ENTITY> data, TableDefinition tableDefinition) {
        writeInternal(tableDefinition, data);
        return this;
    }

    /**
     * 填充数据
     *
     * @param data
     * @param tableDefinition
     * @return
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
     * @param filename
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
     * @return
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
     * @param tableDefinition
     * @param list
     * @param <ENTITY>
     */
    @SuppressWarnings("unchecked")
    private <ENTITY> void writeInternal(TableDefinition tableDefinition, List<ENTITY> list) {
        final Sheet sheet = createSheet(tableDefinition);
        final TableLocation tableLocation = tableDefinition.tableLocation();
        PreparedTableDataHandler preparedTableDataHandler = tableDefinition.preparedTableDataHandler();
        final List finalList = preparedTableDataHandler == null ? list : preparedTableDataHandler.handle(list);
        final ColumnDefinition[] columnDefinitions = tableDefinition.columnDefinitions(workbook, new ColumnDefinitionBuilder(workbook));
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
     * @param tableDefinition
     * @return
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
     * @param sheet
     * @param rowIndex
     * @param startColumnIndex
     * @param title
     * @param columnSize
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
     * @param sheet
     * @param rowIndex
     * @param startColumnIndex
     * @param columnDefinitions
     */
    private void createTableColumnHeader(Sheet sheet, int rowIndex, final int startColumnIndex, ColumnDefinition[] columnDefinitions) {
        Row headerRow = sheet.createRow(rowIndex);
        CellStyle headerCellStyle = DefaultCellStyles.normalCellStyle(workbook);
        for (int i = 0; i < columnDefinitions.length; i++) {
            Cell headerCell = headerRow.createCell(startColumnIndex + i);
            ColumnDefinition<?> columnDefinition = columnDefinitions[i];
            if (columnDefinition == null) {
                headerCell.setCellType(CellType.BLANK);
                continue;
            }
            headerCell.setCellStyle(headerCellStyle);
            if (columnDefinition.getHeader() != null) {
                headerCell.setCellType(CellType.STRING);
                headerCell.setCellValue(columnDefinition.getHeader());
            } else {
                headerCell.setCellType(CellType.BLANK);
            }
        }
    }

    /**
     * 创建表内容
     *
     * @param sheet
     * @param rowIndex
     * @param startColumnIndex
     * @param columnDefinitions
     * @param list
     * @param <ENTITY>
     */
    private <ENTITY> void createTableBody(Sheet sheet, int rowIndex, final int startColumnIndex, ColumnDefinition[] columnDefinitions, List<ENTITY> list) {
        // 构建列单元格风格
        final CellStyle[] columnCellStyles = Stream
                .of(columnDefinitions)
                .map(this::configCellStyle)
                .toArray(CellStyle[]::new);
        // 渲染单元格
        for (int i = 0; i < list.size(); i++) {
            ENTITY entity = list.get(i);
            Row row = sheet.createRow(rowIndex + i);
            for (int j = 0; j < columnDefinitions.length; j++) {
                Cell cell = row.createCell(startColumnIndex + j);
                ColumnDefinition<?> columnDefinition = columnDefinitions[j];
                if (columnDefinition == null || columnDefinition instanceof BlankColumnDefinition) {
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
                cell.setCellStyle(columnCellStyles[j]);
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
     * @param columnDefinition
     * @return
     */
    private CellStyle configCellStyle(ColumnDefinition<?> columnDefinition) {
        CellStyle cellStyle = DefaultCellStyles.normalCellStyle(workbook);
        cellStyle = columnDefinition.cellStyleProvider == null ? cellStyle : columnDefinition.cellStyleProvider.provide(workbook, cellStyle);
        if (columnDefinition.format != null) {
            cellStyle.setDataFormat(workbook.createDataFormat().getFormat(columnDefinition.format));
        }
        if (columnDefinition.alignment != null) {
            cellStyle.setAlignment(columnDefinition.alignment.getFirstValue());
            cellStyle.setVerticalAlignment(columnDefinition.alignment.getSecondValue());
        }
        return cellStyle;
    }
}
