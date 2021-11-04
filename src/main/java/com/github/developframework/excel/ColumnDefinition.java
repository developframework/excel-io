package com.github.developframework.excel;

import develop.toolkit.base.struct.TwoValues;
import lombok.Getter;
import org.apache.poi.ss.usermodel.*;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.function.BiFunction;

/**
 * 列定义
 *
 * @param <CELL_TYPE> 单元格类型
 * @param <FIELD>     装填字段类型
 */
@SuppressWarnings("unused")
@Getter
public abstract class ColumnDefinition<CELL_TYPE, FIELD> {

    protected Workbook workbook;

    protected String header;

    protected CellStyleProvider cellStyleProvider;

    protected String field;

    protected Integer columnWidth;

    protected String format;

    protected TwoValues<HorizontalAlignment, VerticalAlignment> alignment;

    protected DataFormatter dataFormatter;

    protected BiFunction<Object, FIELD, CELL_TYPE> writeConvertFunction;

    protected BiFunction<Object, CELL_TYPE, FIELD> readConvertFunction;

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
    @SuppressWarnings("unchecked")
    protected final void writeIntoCell(Object entity, Cell cell, Object fieldValue) {
        CELL_TYPE convertValue = writeConvertValue(entity, (FIELD) fieldValue);
        if (convertValue != null) {
            setCellValue(cell, convertValue);
        }
    }

    /**
     * 读取单元格值
     *
     * @param entity 实体
     * @param cell   单元格
     * @return 单元格值
     */
    protected final <T> Object readOutCell(Object entity, Cell cell, Class<T> clazz) {
        CELL_TYPE cellValue = getCellValue(cell);
        return readConvertValue(entity, cellValue, clazz);
    }

    /**
     * 该列的CellType
     *
     * @return 单元格类型
     */
    protected abstract CellType getColumnCellType();

    /**
     * 设置单元格值
     *
     * @param cell         单元格
     * @param convertValue 转化值
     */
    protected abstract void setCellValue(Cell cell, CELL_TYPE convertValue);


    /**
     * 写入转化值
     *
     * @param entity     实体
     * @param fieldValue 字段值
     * @return 转化值
     */
    protected abstract CELL_TYPE writeConvertValue(Object entity, FIELD fieldValue);

    /**
     * 读取单元格值
     *
     * @param cell 单元格
     * @return 单元格值
     */
    protected abstract CELL_TYPE getCellValue(Cell cell);

    /**
     * 列宽
     *
     * @param columnWidth 列宽
     */
    public ColumnDefinition<CELL_TYPE, FIELD> columnWidth(int columnWidth) {
        this.columnWidth = columnWidth;
        return this;
    }

    /**
     * 设置单元格风格
     *
     * @param cellStyleProvider 单元格风格
     */
    @SuppressWarnings("unused")
    public ColumnDefinition<CELL_TYPE, FIELD> style(CellStyleProvider cellStyleProvider) {
        this.cellStyleProvider = cellStyleProvider;
        return this;
    }

    /**
     * 设置格式
     *
     * @param format 格式
     */
    @SuppressWarnings("unused")
    public ColumnDefinition<CELL_TYPE, FIELD> format(String format) {
        this.format = format;
        return this;
    }

    /**
     * @param horizontalAlignment 水平对齐
     * @param verticalAlignment   垂直对齐
     */
    @SuppressWarnings("unused")
    public ColumnDefinition<CELL_TYPE, FIELD> alignment(HorizontalAlignment horizontalAlignment, VerticalAlignment verticalAlignment) {
        this.alignment = TwoValues.of(horizontalAlignment, verticalAlignment);
        return this;
    }

    /**
     * 读取转化值
     *
     * @param entity     实体
     * @param cellValue  单元格值
     * @param fieldClass 字段类型
     * @return 字段值
     */
    private Object readConvertValue(Object entity, CELL_TYPE cellValue, Class<?> fieldClass) {
        Object convertValue;
        if (readConvertFunction != null) {
            convertValue = readConvertFunction.apply(entity, cellValue);
        } else {
            convertValue = cellValue;
        }
        if (convertValue == null) {
            return null;
        } else if (fieldClass == convertValue.getClass()) {
            return convertValue;
        } else if (fieldClass == String.class) {
            return convertValue.toString();
        } else if (fieldClass == Integer.class || fieldClass == int.class) {
            return new BigDecimal(convertValue.toString()).intValue();
        } else if (fieldClass == Long.class || fieldClass == long.class) {
            return new BigDecimal(convertValue.toString()).longValue();
        } else if (fieldClass == Boolean.class || fieldClass == boolean.class) {
            return Boolean.valueOf(convertValue.toString());
        } else if (fieldClass == BigDecimal.class) {
            return new BigDecimal(convertValue.toString());
        } else if (fieldClass == Float.class || fieldClass == float.class) {
            return new BigDecimal(convertValue.toString()).floatValue();
        } else if (fieldClass == Double.class || fieldClass == double.class) {
            return new BigDecimal(convertValue.toString()).doubleValue();
        } else if (fieldClass == LocalDateTime.class) {
            return LocalDateTime.parse(convertValue.toString(), format == null ? DateTimeFormatter.ISO_LOCAL_DATE_TIME : DateTimeFormatter.ofPattern(format));
        } else if (fieldClass == LocalDate.class) {
            return LocalDate.parse(convertValue.toString(), format == null ? DateTimeFormatter.ISO_LOCAL_DATE : DateTimeFormatter.ofPattern(format));
        } else if (fieldClass == LocalTime.class) {
            return LocalTime.parse(convertValue.toString(), format == null ? DateTimeFormatter.ISO_LOCAL_TIME : DateTimeFormatter.ofPattern(format));
        } else {
            throw new IllegalArgumentException("can not convert from \"java.lang.String\" to \"" + fieldClass.getName() + "\"");
        }
    }
}
