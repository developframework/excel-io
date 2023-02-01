package com.github.developframework.excel;

import com.github.developframework.excel.column.BlankColumnDefinition;
import com.github.developframework.excel.column.ColumnDefinitionBuilder;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;

/**
 * Excel读取器
 *
 * @author qiushui on 2019-05-18.
 */
@Slf4j
public class ExcelReader extends ExcelProcessor {

    protected ExcelReader(Workbook workbook) {
        super(workbook);
    }

    /**
     * 读取表格内容
     *
     * @param entityClass     实体类型
     * @param tableDefinition 表定义
     * @param <ENTITY>        实体类泛型
     * @return 实体列表
     */
    @SuppressWarnings("unused")
    public <ENTITY> List<ENTITY> read(Class<ENTITY> entityClass, TableDefinition<ENTITY> tableDefinition) {
        return read(entityClass, Integer.MAX_VALUE, tableDefinition);
    }

    /**
     * 读取表格内容
     *
     * @param entityClass     实体类型
     * @param readSize        读取数量
     * @param tableDefinition 表定义
     * @param <ENTITY>        实体类泛型
     * @return 实体列表
     */
    public <ENTITY> List<ENTITY> read(Class<ENTITY> entityClass, Integer readSize, TableDefinition<ENTITY> tableDefinition) {
        final TableInfo tableInfo = tableDefinition.tableInfo();
        // 获取工作表
        Sheet sheet = getSheet(workbook, tableInfo);

        // 表格位置
        TableLocation tableLocation = tableInfo.tableLocation;
        final int totalSize = sheet.getLastRowNum() + 1 - tableLocation.getRow() - (tableInfo.hasTitle ? 1 : 0) - (tableInfo.hasColumnHeader ? 1 : 0);
        final int startColumnIndex = tableLocation.getColumn();
        int rowIndex = tableLocation.getRow() + (tableInfo.hasTitle ? 1 : 0) + (tableInfo.hasColumnHeader ? 1 : 0);

        // 列定义
        final ColumnDefinition<ENTITY>[] columnDefinitions = tableDefinition.columnDefinitions(workbook, new ColumnDefinitionBuilder<>(workbook));
        final int size = Math.min(readSize, totalSize);
        final List<ENTITY> list = new LinkedList<>();
        for (int i = 0; i < size; i++) {
            Row row = sheet.getRow(rowIndex + i);
            ColumnDefinition<ENTITY> columnDefinition = null;
            try {
                final ENTITY entity = entityClass.getConstructor().newInstance();
                for (int j = 0; j < columnDefinitions.length; j++) {
                    columnDefinition = columnDefinitions[j];
                    if (columnDefinition == null || columnDefinition instanceof BlankColumnDefinition) {
                        continue;
                    }
                    Cell cell = row.getCell(startColumnIndex + j);
                    if (cell != null) {
                        // 读取单元格值装填到实体
                        columnDefinition.readOutCell(workbook, cell, entity);
                    }
                }
                tableDefinition.each(entity);
                list.add(entity);
            } catch (Exception e) {
                assert columnDefinition != null;
                log.error("row {} column {}", row.getRowNum(), columnDefinition.getColumnInfo().field);
                throw new RuntimeException(e);
            }
        }
        return new ArrayList<>(list);
    }


    private Sheet getSheet(Workbook workbook, TableInfo tableInfo) {
        if (tableInfo.sheetName != null) {
            return workbook.getSheet(tableInfo.sheetName);
        } else {
            return workbook.getSheetAt(tableInfo.sheet);
        }
    }
}
