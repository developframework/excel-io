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

    protected DataFormatter dataFormatter;

    public ColumnDefinition(Workbook workbook, String field, String header) {
        this.workbook = workbook;
        this.field = field;
        this.header = header;
        this.dataFormatter = new DataFormatter();
    }

    /**
     * 值写入单元格
     *
     * @param entity     实体
     * @param cell       单元格
     * @param fieldValue 字段值
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
     * @param entity 实体
     * @param cell 单元格
     * @return 单元格值
     */
    protected <T> Object readOutCell(Object entity, Cell cell, Class<T> clazz) {
        TYPE cellValue = getCellValue(cell);
        return readConvertValue(entity, cellValue, clazz);
    }

    /**
     * 该列的CellType
     *
     * @return 单元格类型
     */
    @SuppressWarnings("unused")
    protected abstract CellType getColumnCellType();

    /**
     * 设置单元格值
     *
     * @param cell 单元格
     * @param convertValue 转化值
     */
    protected abstract void setCellValue(Cell cell, TYPE convertValue);


    /**
     * 写入转化值
     *
     * @param entity 实体
     * @param fieldValue 字段值
     * @return 转化值
     */
    protected abstract TYPE writeConvertValue(Object entity, Object fieldValue);

    /**
     * 读取单元格值
     *
     * @param cell 单元格
     * @return 单元格值
     */
    protected abstract TYPE getCellValue(Cell cell);

    /**
     * 读取转化值
     *
     * @param entity 实体
     * @param cellValue 单元格值
     * @param fieldClass 字段类型
     * @return 字段值
     */
    protected abstract <T> Object readConvertValue(Object entity, TYPE cellValue, Class<T> fieldClass);

    /**
     * 列宽
     *
     * @param columnWidth 列宽
     */
    @SuppressWarnings("unused")
    public ColumnDefinition<TYPE> columnWidth(int columnWidth) {
        this.columnWidth = columnWidth;
        return this;
    }

    /**
     * 设置单元格风格
     *
     * @param cellStyleProvider 单元格风格
     */
    @SuppressWarnings("unused")
    public ColumnDefinition<TYPE> style(CellStyleProvider cellStyleProvider) {
        this.cellStyleProvider = cellStyleProvider;
        return this;
    }

    /**
     * 设置格式
     *
     * @param format 格式
     */
    @SuppressWarnings("unused")
    public ColumnDefinition<TYPE> format(String format) {
        this.format = format;
        return this;
    }

    /**
     * @param horizontalAlignment 水平对齐
     * @param verticalAlignment   垂直对齐
     */
    @SuppressWarnings("unused")
    public ColumnDefinition<TYPE> alignment(HorizontalAlignment horizontalAlignment, VerticalAlignment verticalAlignment) {
        this.alignment = TwoValues.of(horizontalAlignment, verticalAlignment);
        return this;
    }
}
