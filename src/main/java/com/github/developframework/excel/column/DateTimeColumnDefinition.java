package com.github.developframework.excel.column;

import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Workbook;

import java.lang.reflect.Field;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.ZoneId;
import java.util.Calendar;
import java.util.Date;

/**
 * 日期时间列定义
 *
 * @author qiushui on 2018-10-10.
 * @since 0.1
 */
public class DateTimeColumnDefinition extends BasicColumnDefinition {

    public DateTimeColumnDefinition(Workbook workbook, String field) {
        super(workbook, field);
        this.cellType = CellType.NUMERIC;
        DataFormat dataFormat = workbook.createDataFormat();
        this.cellStyle.setDataFormat(dataFormat.getFormat("yyyy-MM-dd HH:mm:ss"));
    }

    @Override
    public void dealFillData(Cell cell, Object value) {
        Class<?> valueClass = value.getClass();
        if (LocalDate.class.isAssignableFrom(valueClass)) {
            LocalTime localTime = LocalTime.of(0, 0, 0, 0);
            LocalDateTime localDateTime = LocalDateTime.of((LocalDate) value, localTime);
            cell.setCellValue(Date.from(localDateTime.atZone(ZoneId.systemDefault()).toInstant()));
        } else if (LocalDateTime.class.isAssignableFrom(valueClass)) {
            cell.setCellValue(Date.from(((LocalDateTime) value).atZone(ZoneId.systemDefault()).toInstant()));
        } else if (Date.class.isAssignableFrom(valueClass)) {
            cell.setCellValue((Date) value);
        } else if (Calendar.class.isAssignableFrom(valueClass)) {
            cell.setCellValue((Calendar) value);
        }
    }

    @Override
    public void dealReadData(Cell cell, Object instance) {
        Class<?> instanceClass = instance.getClass();
        Field field = FieldUtils.getField(instanceClass, fieldName, true);
        if (field.getType() == Date.class) {
            try {
                Date value = cell.getDateCellValue();
                Object object = readColumnValueConverter.map(converter -> converter.convert(instance, value)).orElse(value);
                FieldUtils.writeDeclaredField(instance, fieldName, object, true);
            } catch (IllegalAccessException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 设置格式
     *
     * @param pattern
     * @return
     */
    public ColumnDefinition pattern(String pattern) {
        DataFormat dataFormat = workbook.createDataFormat();
        this.cellStyle.setDataFormat(dataFormat.getFormat(pattern));
        return this;
    }
}
