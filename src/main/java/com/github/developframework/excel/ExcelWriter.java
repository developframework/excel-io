package com.github.developframework.excel;

import com.github.developframework.excel.column.ColumnDefinition;
import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;

import java.io.IOException;
import java.io.OutputStream;
import java.util.Arrays;
import java.util.List;

/**
 * Excel写出器
 *
 * @author qiushui on 2018-10-09.
 * @since 0.1
 */
public class ExcelWriter extends ExcelProcessor {

    private OutputStream outputStream;

    public ExcelWriter(Workbook workbook, OutputStream outputStream) {
        super(workbook);
        this.outputStream = outputStream;
    }

    /**
     * 填充数据
     *
     * @param data
     * @param tableDefinition
     * @return
     */
    public <T> ExcelWriter fillData(List<T> data, TableDefinition tableDefinition) {
        dealFillData(workbook, data, tableDefinition);
        return this;
    }

    /**
     * 填充数据
     *
     * @param data
     * @param tableDefinition
     * @return
     */
    public <T> ExcelWriter fillData(T[] data, TableDefinition tableDefinition) {
        dealFillData(workbook, Arrays.asList(data), tableDefinition);
        return this;
    }

    /**
     * 写出
     */
    public void write() {
        try {
            IOUtils.write(workbook, outputStream);
            workbook.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 填充数据
     *
     * @param workbook
     * @param list
     * @param tableDefinition
     */
    private <T> void dealFillData(Workbook workbook, List<T> list, TableDefinition tableDefinition) {
        Sheet sheet = getSheet(workbook, tableDefinition);
        int rowIndex = tableDefinition.row();
        int columnIndex;
        ColumnDefinition[] columnDefinitions = tableDefinition.columnDefinitions(workbook);
        // 填充表头
        if(tableDefinition.hasHeader()) {
            Row headerRow = sheet.createRow(rowIndex++);
            columnIndex = tableDefinition.column();
            CellStyle headerCellStyle = workbook.createCellStyle();
            tableDefinition.tableHeaderCellStyle(workbook, headerCellStyle);
            for (int i = 0; i < columnDefinitions.length; i++) {
                Cell headerCell = headerRow.createCell(columnIndex + i);
                headerCell.setCellStyle(headerCellStyle);
                headerCell.setCellType(CellType.STRING);
                if (columnDefinitions[i] == null) {
                    continue;
                }
                headerCell.setCellValue(columnDefinitions[i].getHeader());
            }
        }

        // 填充表内容
        for (int i = 0; i < list.size(); i++) {
            T item = list.get(i);
            Row row = sheet.createRow(rowIndex + i);
            columnIndex = tableDefinition.column();
            for (int j = 0; j < columnDefinitions.length; j++) {
                ColumnDefinition columnDefinition = columnDefinitions[j];
                if(columnDefinition == null) {
                    continue;
                }
                Cell cell = row.createCell(columnIndex + j);
                cell.setCellType(columnDefinition.getCellType());
                cell.setCellStyle(columnDefinition.getCellStyle());
                try {
                    Object value = FieldUtils.readDeclaredField(item, columnDefinition.getFieldName(), true);
                    columnDefinition.fillData(cell, value);
                } catch (IllegalAccessException e) {
                    throw new RuntimeException(e);
                }
            }
        }

        // 自动列宽
        for (int i = 0; i < columnDefinitions.length; i++) {
            sheet.autoSizeColumn(i);
        }
    }

    private Sheet getSheet(Workbook workbook, TableDefinition tableDefinition) {
        String sheetName = tableDefinition.sheetName() == null ? ("sheet " + (workbook.getNumberOfSheets() + 1)) : tableDefinition.sheetName();
        return workbook.createSheet(sheetName);
    }
}
