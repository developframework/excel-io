package com.github.developframework.excel.column;

import lombok.Getter;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;

import java.util.Map;

/**
 * @author qiushui on 2018-10-10.
 * @since 0.1
 */
@Getter
public abstract class ColumnDefinition {

    protected String header;

    protected CellStyle cellStyle;

    protected CellType cellType;

    protected String fieldName;

    protected Integer maxLength;

    /**
     * 加边框
     *
     * @param cellStyle
     */
    protected void borderThin(CellStyle cellStyle) {
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
    }

    /**
     * 填充数据
     *
     * @param cell
     * @param value
     */
    public final void fillData(Cell cell, Object value) {
        if (value == null) {
            cell.setCellType(CellType.BLANK);
        } else {
            dealFillData(cell, value);
        }
    }

    /**
     * 读取数据
     *
     * @param cell
     * @param instance
     */
    @SuppressWarnings("unchecked")
    public final void readData(Cell cell, Object instance) {
        if(instance != null) {
            Class<?> instanceClass = instance.getClass();
            if(Map.class.isAssignableFrom(instanceClass)) {
                String value = cell.getStringCellValue();
                ((Map<String, Object>) instance).put(fieldName, value);
            } else if(cell.getCellTypeEnum() != CellType.BLANK){
                dealReadData(cell, instance);
            }
        }
    }

    /**
     * 数据填充
     *
     * @param cell
     * @param value
     */
    public abstract void dealFillData(Cell cell, Object value);

    /**
     * 数据读取
     *
     * @param cell
     * @param instance
     */
    public abstract void dealReadData(Cell cell, Object instance);
}
