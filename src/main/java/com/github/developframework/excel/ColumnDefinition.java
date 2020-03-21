package com.github.developframework.excel;

import develop.toolkit.base.struct.TwoValues;
import lombok.Getter;
import org.apache.poi.ss.usermodel.*;

/**
 * 列定义
 *
 * @author qiushui on 2018-10-10.
 */
@Getter
public abstract class ColumnDefinition<TYPE> {

    protected Workbook workbook;

    protected String header;

    protected CellStyleProvider cellStyleProvider;

    protected String field;

    protected Integer columnWidth;

    protected String format;

    protected TwoValues<HorizontalAlignment, VerticalAlignment> alignment;

    public ColumnDefinition(Workbook workbook, String field, String header) {
        this.workbook = workbook;
        this.field = field;
        this.header = header;
    }

    /**
     * 值写入单元格
     *
     * @param entity
     * @param cell
     * @param fieldValue
     */
    protected void writeIntoCell(Object entity, Cell cell, Object fieldValue) {
        TYPE convertValue = writeConvertValue(entity, fieldValue);
        if (convertValue != null) {
            setCellValue(cell, convertValue);
        }
    }

    /**
     * 读取单元格值
     *
     * @param entity
     * @param cell
     * @return
     */
    protected <T> Object readOutCell(Object entity, Cell cell, Class<T> clazz) {
        TYPE cellValue = getCellValue(cell);
        return readConvertValue(entity, cellValue, clazz);
    }

    /**
     * 该列的CellType
     *
     * @return
     */
    protected abstract CellType getColumnCellType();

    /**
     * 设置单元格值
     *
     * @param cell
     * @param convertValue
     */
    protected abstract void setCellValue(Cell cell, TYPE convertValue);


    /**
     * 写入转化值
     *
     * @param entity
     * @param fieldValue
     * @return
     */
    protected abstract TYPE writeConvertValue(Object entity, Object fieldValue);

    /**
     * 读取单元格值
     *
     * @param cell
     * @return
     */
    protected abstract TYPE getCellValue(Cell cell);

    /**
     * 读取转化值
     *
     * @param cellValue
     * @return
     */
    protected abstract <T> Object readConvertValue(Object entity, TYPE cellValue, Class<T> fieldClass);

    /**
     * 列宽
     *
     * @param columnWidth
     * @return
     */
    public ColumnDefinition columnWidth(int columnWidth) {
        this.columnWidth = columnWidth;
        return this;
    }

    /**
     * 设置单元格风格
     *
     * @param cellStyleProvider
     * @return
     */
    public ColumnDefinition style(CellStyleProvider cellStyleProvider) {
        this.cellStyleProvider = cellStyleProvider;
        return this;
    }

    /**
     * 设置格式
     *
     * @param format
     * @return
     */
    public ColumnDefinition format(String format) {
        this.format = format;
        return this;
    }

    /**
     * @param horizontalAlignment
     * @param verticalAlignment
     * @return
     */
    public ColumnDefinition alignment(HorizontalAlignment horizontalAlignment, VerticalAlignment verticalAlignment) {
        this.alignment = TwoValues.of(horizontalAlignment, verticalAlignment);
        return this;
    }
}
