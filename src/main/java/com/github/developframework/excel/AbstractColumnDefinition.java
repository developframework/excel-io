package com.github.developframework.excel;

import com.github.developframework.expression.ExpressionUtils;
import lombok.SneakyThrows;
import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Workbook;

import java.math.BigDecimal;
import java.time.*;
import java.util.Date;
import java.util.function.BiFunction;

/**
 * 列定义
 *
 * @param <ENTITY> 实体类型
 * @param <FIELD>  装填字段类型
 */
@SuppressWarnings("unused")
public abstract class AbstractColumnDefinition<ENTITY, FIELD> implements ColumnDefinition<ENTITY> {

    protected ColumnInfo columnInfo;
    protected DataFormatter dataFormatter;
    protected BiFunction<ENTITY, FIELD, Object> writeConvertFunction;
    protected BiFunction<ENTITY, Object, FIELD> readConvertFunction;

    protected BiFunction<Cell, Object, String> cellStyleKeyProvider;

    public AbstractColumnDefinition(String field, String header) {
        this.columnInfo = new ColumnInfo(field, header);
        this.dataFormatter = new DataFormatter();
    }

    /**
     * 值写入单元格
     *
     * @param workbook 工作区
     * @param cell     单元格
     * @param entity   实体
     */
    @Override
    public final Object writeIntoCell(Workbook workbook, Cell cell, ENTITY entity) {
        final FIELD fieldValue = getEntityValue(entity);
        final Object convertValue = writeConvertFunction == null ? fieldValue : writeConvertFunction.apply(entity, fieldValue);
        setCellValue(cell, convertValue);
        return convertValue;
    }

    /**
     * 读取单元格值
     *
     * @param workbook 工作区
     * @param cell     单元格
     * @param entity   实体
     */
    @Override
    @SuppressWarnings("unchecked")
    public final void readOutCell(Workbook workbook, Cell cell, ENTITY entity) {
        final Object cellValue = getCellValue(cell);
        final FIELD convertValue = readConvertFunction == null ? (FIELD) cellValue : readConvertFunction.apply(entity, cellValue);
        if (convertValue != null) {
            setEntityValue(entity, convertValue);
        }
    }

    /**
     * 设置单元格值
     *
     * @param cell         单元格
     * @param convertValue 转化值
     */
    protected void setCellValue(Cell cell, Object convertValue) {
        if (convertValue == null) {
            cell.setBlank();
        } else {
            final Class<?> clazz = convertValue.getClass();
            if (clazz == String.class) {
                cell.setCellValue((String) convertValue);
            } else if (clazz == Integer.class || clazz == Integer.TYPE) {
                cell.setCellValue(((Integer) convertValue).doubleValue());
            } else if (clazz == Float.class || clazz == Float.TYPE) {
                cell.setCellValue(((Float) convertValue).doubleValue());
            } else if (clazz == Double.class || clazz == Double.TYPE) {
                cell.setCellValue((Double) convertValue);
            } else if (clazz == BigDecimal.class) {
                cell.setCellValue(((BigDecimal) convertValue).doubleValue());
            } else if (clazz == Boolean.class || clazz == Boolean.TYPE) {
                cell.setCellValue((Boolean) convertValue);
            } else if (clazz == LocalDateTime.class) {
                cell.setCellValue(Date.from(((LocalDateTime) convertValue).atZone(ZoneId.systemDefault()).toInstant()));
            } else if (clazz == ZonedDateTime.class) {
                cell.setCellValue(Date.from(((ZonedDateTime) convertValue).toInstant()));
            } else if (clazz == LocalDate.class || clazz == LocalTime.class) {
                cell.setCellValue(convertValue.toString());
            } else if (clazz == java.util.Date.class) {
                cell.setCellValue((java.util.Date) convertValue);
            }
        }
    }

    /**
     * 读取单元格值
     *
     * @param cell 单元格
     * @return 单元格值
     */
    protected abstract Object getCellValue(Cell cell);

    /**
     * 赋值给实体
     *
     * @param entity     实体
     * @param fieldValue 字段值
     */
    @SneakyThrows(IllegalAccessException.class)
    protected void setEntityValue(ENTITY entity, FIELD fieldValue) {
        FieldUtils.writeDeclaredField(entity, columnInfo.field, fieldValue, true);
    }

    /**
     * 读取实体值
     *
     * @param entity 实体
     * @return 实体字段值
     */
    @SuppressWarnings("unchecked")
    protected FIELD getEntityValue(ENTITY entity) {
        return (FIELD) ExpressionUtils.getValue(entity, columnInfo.field);
    }

    /**
     * 列宽
     *
     * @param columnWidth 列宽
     */
    public final AbstractColumnDefinition<ENTITY, FIELD> columnWidth(int columnWidth) {
        this.columnInfo.columnWidth = columnWidth;
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
    private Object readConvertValue(ENTITY entity, Object cellValue, Class<?> fieldClass) {
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
        } else {
            throw new IllegalArgumentException("can not convert from \"java.lang.String\" to \"" + fieldClass.getName() + "\"");
        }
    }

    @Override
    public void configureCellStyle(Cell cell, CellStyleManager cellStyleManager, Object value) {
        final String key;
        if (cellStyleKeyProvider != null) {
            key = cellStyleKeyProvider.apply(cell, value);
        } else {
            key = determineCellStyleKey(cell, value);
        }
        cell.setCellStyle(cellStyleManager.getCellStyle(key));
    }

    /**
     * 决定单元格格式键
     */
    protected String determineCellStyleKey(Cell cell, Object value) {
        final Class<?> valueClass = value.getClass();
        if (valueClass == LocalDateTime.class || valueClass == ZonedDateTime.class || valueClass == java.util.Date.class) {
            return CellStyleManager.STYLE_NORMAL_DATETIME;
        }
        return CellStyleManager.STYLE_NORMAL;
    }

    public AbstractColumnDefinition<ENTITY, FIELD> writeConvert(BiFunction<ENTITY, FIELD, Object> writeConvertFunction) {
        this.writeConvertFunction = writeConvertFunction;
        return this;
    }

    public AbstractColumnDefinition<ENTITY, FIELD> readConvert(BiFunction<ENTITY, Object, FIELD> readConvertFunction) {
        this.readConvertFunction = readConvertFunction;
        return this;
    }

    public AbstractColumnDefinition<ENTITY, FIELD> cellStyleKey(BiFunction<Cell, Object, String> cellStyleKeyProvider) {
        this.cellStyleKeyProvider = cellStyleKeyProvider;
        return this;
    }
}
