package com.github.developframework.excel;

import com.github.developframework.excel.column.BlankColumnDefinition;
import com.github.developframework.excel.column.ColumnDefinitionBuilder;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;

/**
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
     * @param entityClass
     * @param tableDefinition
     * @param <ENTITY>
     * @return
     */
    public <ENTITY> List<ENTITY> read(Class<ENTITY> entityClass, TableDefinition tableDefinition) {
        return read(entityClass, null, tableDefinition);
    }

    /**
     * 读取表格内容
     *
     * @param entityClass
     * @param readSize
     * @param tableDefinition
     * @param <ENTITY>
     * @return
     */
    public <ENTITY> List<ENTITY> read(Class<ENTITY> entityClass, Integer readSize, TableDefinition tableDefinition) {
        Sheet sheet = getSheet(workbook, tableDefinition);
        TableLocation tableLocation = tableDefinition.tableLocation();
        final int totalSize = sheet.getLastRowNum() + 1 - tableLocation.getRow() - (tableDefinition.hasTitle() ? 1 : 0) - (tableDefinition.hasColumnHeader() ? 1 : 0);
        final int startColumnIndex = tableLocation.getColumn();
        int rowIndex = tableLocation.getRow() + (tableDefinition.hasTitle() ? 1 : 0) + (tableDefinition.hasColumnHeader() ? 1 : 0);
        ColumnDefinition[] columnDefinitions = tableDefinition.columnDefinitions(workbook, new ColumnDefinitionBuilder(workbook));
        List<ENTITY> list = new LinkedList<>();
        final int size = readSize != null && readSize < totalSize ? readSize : totalSize;
        for (int i = 0; i < size; i++) {
            Row row = sheet.getRow(rowIndex + i);
            ColumnDefinition<?> columnDefinition = null;
            try {
                ENTITY entity = entityClass.getConstructor().newInstance();
                for (int j = 0; j < columnDefinitions.length; j++) {
                    columnDefinition = columnDefinitions[j];
                    if (columnDefinition == null || columnDefinition instanceof BlankColumnDefinition) {
                        continue;
                    }
                    Cell cell = row.getCell(startColumnIndex + j);
                    if (cell != null) {
                        Field field = FieldUtils.getDeclaredField(entityClass, columnDefinition.field, true);
                        Object value = columnDefinition.readOutCell(entity, cell, field.getType());
                        FieldUtils.writeDeclaredField(entity, columnDefinition.field, value, true);
                    }
                }
                list.add(entity);
            } catch (Exception e) {
                assert columnDefinition != null;
                log.error("row {} column {}", row.getRowNum(), columnDefinition.field);
                throw new RuntimeException(e);
            }
        }
        return new ArrayList<>(list);
    }



    private Sheet getSheet(Workbook workbook, TableDefinition tableDefinition) {
        if (tableDefinition.sheet() != null) {
            return workbook.getSheetAt(tableDefinition.sheet());
        } else if(tableDefinition.sheetName() != null) {
            return workbook.getSheet(tableDefinition.sheetName());
        } else {
            return workbook.getSheetAt(0);
        }
    }
}
