package com.github.developframework.excel;

import com.github.developframework.excel.styles.DefaultCellStyles;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.HashMap;
import java.util.Map;

/**
 * 单元格样式管理器
 *
 * @author qiushui on 2022-06-29.
 */
public class CellStyleManager {

    public static final String STYLE_NORMAL = "normal";
    public static final String STYLE_NORMAL_DATETIME = "normalDateTime";
    public static final String STYLE_NORMAL_NUMBER = "normalNumber";

    private final Map<String, CellStyle> cellStyleMap = new HashMap<>();

    public CellStyleManager(Workbook workbook, TableDefinition<?> tableDefinition) {
        cellStyleMap.put(STYLE_NORMAL, DefaultCellStyles.normalCellStyle(workbook));
        cellStyleMap.put(STYLE_NORMAL_DATETIME, DefaultCellStyles.normalDateTimeCellStyle(workbook));
        cellStyleMap.put(STYLE_NORMAL_NUMBER, DefaultCellStyles.numberCellStyle(workbook));
        cellStyleMap.putAll(tableDefinition.customCellStyles(workbook));
    }

    public CellStyle getCellStyle(String key) {
        return cellStyleMap.getOrDefault(key, cellStyleMap.get("normal"));
    }
}
