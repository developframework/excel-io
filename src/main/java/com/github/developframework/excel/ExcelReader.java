package com.github.developframework.excel;

import com.github.developframework.excel.column.ColumnDefinition;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;

/**
 * Excel读取器
 *
 * @author qiushui on 2018-10-10.
 * @since 0.1
 */
public class ExcelReader extends ExcelProcessor {

    public ExcelReader(Workbook workbook) {
        super(workbook);
    }

    /**
     * 读取
     *
     * @param clazz
     * @param tableDefinition
     * @param readSize
     * @return
     */
    public <T> List<T> read(Class<T> clazz, Integer readSize, TableDefinition tableDefinition) {
        List<T> list = new LinkedList<>();
        Sheet sheet = getSheet(workbook, tableDefinition);
        int lastRowNum = sheet.getLastRowNum();
        int size = (readSize == null || readSize >= lastRowNum ? lastRowNum : readSize) - tableDefinition.row();
        int rowIndex = tableDefinition.row() + (tableDefinition.hasHeader() ? 1 : 0);
        int columnIndex = tableDefinition.column();
        ColumnDefinition[] columnDefinitions = tableDefinition.columnDefinitions(workbook);
        for (int i = 0; i < size; i++) {
            Row row = sheet.getRow(rowIndex + i);
            try {
                T item = clazz.getConstructor().newInstance();
                for (int j = 0; j < columnDefinitions.length; j++) {
                    ColumnDefinition columnDefinition = columnDefinitions[j];
                    if(columnDefinition == null) {
                        continue;
                    }
                    Cell cell = row.getCell(columnIndex + j);
                    if (cell != null) {
                        cell.setCellStyle(columnDefinition.getCellStyle());
                        cell.setCellType(columnDefinition.getCellType());
                        columnDefinition.readData(cell, item);
                    }
                }
                list.add(item);
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
        }
        return new ArrayList<>(list);
    }

    /**
     * 读取
     *
     * @param clazz
     * @param tableDefinition
     * @param <T>
     * @return
     */
    public <T> List<T> read(Class<T> clazz, TableDefinition tableDefinition) {
        return read(clazz, null, tableDefinition);
    }

    private Sheet getSheet(Workbook workbook, TableDefinition tableDefinition) {
        if (tableDefinition.sheet() != null) {
            return workbook.getSheetAt(tableDefinition.sheet());
        } else if(tableDefinition.sheetName() != null) {
            return workbook.getSheet(tableDefinition.sheetName());
        } else {
            throw new RuntimeException("sheet name and index is null");
        }
    }

    /**
     * 读取并关闭
     * @param clazz
     * @param tableDefinition
     * @param readSize
     * @return
     */
    public <T> List<T> readAndClose(Class<T> clazz, Integer readSize, TableDefinition tableDefinition) {
        List<T> list = read(clazz, readSize, tableDefinition);
        close();
        return list;
    }

    /**
     * 读取并关闭
     *
     * @param clazz
     * @param tableDefinition
     * @param <T>
     * @return
     */
    public <T> List<T> readAndClose(Class<T> clazz, TableDefinition tableDefinition) {
        return read(clazz, null, tableDefinition);
    }

    /**
     * 关闭
     */
    public void close() {
        try {
            workbook.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
