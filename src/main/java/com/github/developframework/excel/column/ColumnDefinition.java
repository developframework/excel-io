package com.github.developframework.excel.column;

import com.github.developframework.excel.styles.DefaultCellStyles;
import lombok.Getter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Map;
import java.util.Optional;

/**
 * @author qiushui on 2018-10-10.
 * @since 0.1
 */
@Getter
public abstract class ColumnDefinition {

    protected Workbook workbook;

    protected String header;

    protected CellStyle cellStyle;

    protected CellType cellType;

    protected String fieldName;

    protected Integer maxLength;

    protected Optional<ColumnValueConverter> readColumnValueConverter = Optional.empty();

    protected Optional<ColumnValueConverter> writeColumnValueConverter = Optional.empty();

    public ColumnDefinition(Workbook workbook) {
        this.workbook = workbook;
        cellStyle = DefaultCellStyles.normalCellStyle(workbook);
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
                Object object = readColumnValueConverter.map(converter -> converter.convert(instance, value)).orElse(value);
                ((Map<String, Object>) instance).put(fieldName, object);
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

    /**
     * 设置列名
     *
     * @param header
     * @return
     */
    public ColumnDefinition header(String header) {
        this.header = header;
        return this;
    }

    /**
     * 设置转换器
     *
     * @param columnValueConverter
     * @return
     */
    public ColumnDefinition readConverter(ColumnValueConverter columnValueConverter) {
        this.readColumnValueConverter = Optional.of(columnValueConverter);
        return this;
    }

    /**
     * 设置转换器
     *
     * @param columnValueConverter
     * @return
     */
    public ColumnDefinition writeConverter(ColumnValueConverter columnValueConverter) {
        this.writeColumnValueConverter = Optional.of(columnValueConverter);
        return this;
    }

    /**
     * 设置最大字数
     *
     * @param maxLength
     * @return
     */
    public ColumnDefinition maxLength(int maxLength) {
        this.maxLength = maxLength;
        return this;
    }

    /**
     * 处理CellStyle
     *
     * @param processor
     * @return
     */
    public ColumnDefinition style(ColumnCellStyleProcessor processor) {
        this.cellStyle = processor.process(cellStyle);
        return this;
    }

    /**
     * 列值转换器
     */
    public interface ColumnValueConverter {

        Object convert(Object data, Object currentValue);
    }

    /**
     * 单元格风格处理器
     */
    public interface ColumnCellStyleProcessor {

        CellStyle process(CellStyle style);
    }
}
