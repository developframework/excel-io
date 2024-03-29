package com.github.developframework.excel;

import com.github.developframework.excel.styles.CellStyleManager;
import com.github.developframework.excel.styles.DefaultCellStyles;
import com.github.developframework.excel.utils.ValueConvertUtils;
import com.github.developframework.expression.ExpressionUtils;
import lombok.Getter;
import lombok.SneakyThrows;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import java.lang.reflect.Field;
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

    @Getter
    protected ColumnInfo columnInfo;
    protected BiFunction<ENTITY, FIELD, Object> writeConvertFunction;
    protected BiFunction<ENTITY, Object, FIELD> readConvertFunction;

    protected CellStyleKeyProvider<ENTITY> cellStyleKeyProvider;

    protected CellCommentInfo cellCommentInfo;

    public AbstractColumnDefinition(String field, String header) {
        this.columnInfo = new ColumnInfo(field, header == null ? field : header);
    }

    /**
     * 值写入单元格
     *
     * @param workbook 工作区
     * @param cell     单元格
     * @param entity   实体
     * @param index    实体索引
     */
    @Override
    public Object writeIntoCell(Workbook workbook, Sheet sheet, Cell cell, ENTITY entity, int index) {
        final FIELD fieldValue = getEntityValue(entity);
        final Object convertValue = writeConvertFunction == null ? fieldValue : writeConvertFunction.apply(entity, fieldValue);
        setCellValue(cell, convertValue);

        if (cellCommentInfo != null) {
            final Object c = ExpressionUtils.getValue(entity, cellCommentInfo.getCommentField());
            final Object a = ExpressionUtils.getValue(entity, cellCommentInfo.getAuthorField());
            if(c != null && a != null) {
                Drawing drawing = sheet.createDrawingPatriarch();
                final ClientAnchor anchor = cellCommentInfo.getAnchorFunction().apply(drawing, cell);
                anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_DONT_RESIZE);
                Comment comment = drawing.createCellComment(anchor);
                comment.setString(new XSSFRichTextString(c.toString()));
                comment.setAuthor(a.toString());
                cell.setCellComment(comment);
            }
        }

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
    public void readOutCell(Workbook workbook, Cell cell, ENTITY entity) {
        final Field field = FieldUtils.getDeclaredField(entity.getClass(), columnInfo.field, true);
        final Object cellValue = getCellValue(cell, field.getType());
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
            } else {
                cell.setCellValue(convertValue.toString());
            }
        }
    }

    /**
     * 读取单元格值
     *
     * @param cell 单元格
     * @return 单元格值
     */
    public Object getCellValue(Cell cell, Class<?> fieldClass) {
        final Object value;
        switch (cell.getCellType()) {
            case STRING:
                value = ValueConvertUtils.stringConvert(cell.getRichStringCellValue().getString().trim(), fieldClass);
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    value = ValueConvertUtils.dateConvert(cell.getDateCellValue(), fieldClass);
                } else {
                    cell.setCellType(CellType.STRING);
                    value = ValueConvertUtils.doubleConvert(cell.getStringCellValue().trim(), fieldClass);
                }
                break;
            case BOOLEAN:
                value = ValueConvertUtils.booleanConvert(cell.getBooleanCellValue(), fieldClass);
                break;
//            case FORMULA:
//                value = cell.getCellFormula();
//                break;
            default:
                value = null;
                break;
        }
        return value;
    }

    /**
     * 赋值给实体
     *
     * @param entity 实体
     * @param value  值
     */
    @SneakyThrows(IllegalAccessException.class)
    protected void setEntityValue(ENTITY entity, Object value) {
        FieldUtils.writeDeclaredField(entity, columnInfo.field, value, true);
    }

    /**
     * 读取实体值
     *
     * @param entity 实体
     * @return 实体字段值
     */
    @SuppressWarnings("unchecked")
    protected FIELD getEntityValue(ENTITY entity) {
        return columnInfo.field == null ? null : (FIELD) ExpressionUtils.getValue(entity, columnInfo.field);
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

    @Override
    public void configureCellStyle(Cell cell, CellStyleManager cellStyleManager, ENTITY entity, Object value) {
        String cellStyleKey = null;
        if (cellStyleKeyProvider != null) {
            cellStyleKey = cellStyleKeyProvider.provideCellStyleKey(cell, entity, value);
        }
        if (StringUtils.isEmpty(cellStyleKey)) {
            cellStyleKey = determineCellStyleKey(cell, value);
        }
        cell.setCellStyle(cellStyleManager.getCellStyle(cellStyleKey));
    }

    /**
     * 根据值决定单元格格式键
     */
    protected String determineCellStyleKey(Cell cell, Object value) {
        if (cell.getCellType() == CellType.NUMERIC) {
            final Class<?> valueClass = value.getClass();
            if (valueClass == LocalDateTime.class || valueClass == ZonedDateTime.class || valueClass == java.util.Date.class) {
                return DefaultCellStyles.STYLE_BODY_DATETIME;
            } else if (Number.class.isAssignableFrom(valueClass)) {
                return DefaultCellStyles.STYLE_BODY_NUMBER;
            }
        }
        return DefaultCellStyles.STYLE_BODY;
    }

    public AbstractColumnDefinition<ENTITY, FIELD> writeConvert(BiFunction<ENTITY, FIELD, Object> writeConvertFunction) {
        this.writeConvertFunction = writeConvertFunction;
        return this;
    }

    public AbstractColumnDefinition<ENTITY, FIELD> readConvert(BiFunction<ENTITY, Object, FIELD> readConvertFunction) {
        this.readConvertFunction = readConvertFunction;
        return this;
    }

    public AbstractColumnDefinition<ENTITY, FIELD> cellStyleKey(CellStyleKeyProvider<ENTITY> cellStyleKeyProvider) {
        this.cellStyleKeyProvider = cellStyleKeyProvider;
        return this;
    }

    public AbstractColumnDefinition<ENTITY, FIELD> comment(String author, String commentField, BiFunction<Drawing, Cell, ClientAnchor> anchorFunction) {
        this.cellCommentInfo = new CellCommentInfo(author, commentField, anchorFunction);
        return this;
    }
}
