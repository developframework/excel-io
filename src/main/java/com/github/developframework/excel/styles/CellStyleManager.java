package com.github.developframework.excel.styles;

import com.github.developframework.excel.TableDefinition;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.HashMap;
import java.util.Map;
import java.util.function.Consumer;

/**
 * 单元格样式管理器
 *
 * @author qiushui on 2022-06-29.
 */
public class CellStyleManager {

    private final Map<String, CellStyle> cellStyleMap = new HashMap<>();

    public CellStyleManager(Workbook workbook, TableDefinition<?> tableDefinition) {
        final Consumer<CellStyle> globalConsumer = tableDefinition.globalCellStylesHandle();
        registerCellStyle(globalConsumer, DefaultCellStyles.STYLE_NORMAL, DefaultCellStyles.normalCellStyle(workbook));
        registerCellStyle(globalConsumer, DefaultCellStyles.STYLE_NORMAL_DATETIME, DefaultCellStyles.normalDateTimeCellStyle(workbook));
        registerCellStyle(globalConsumer, DefaultCellStyles.STYLE_NORMAL_NUMBER, DefaultCellStyles.normalNumberCellStyle(workbook));
        registerCellStyle(globalConsumer, DefaultCellStyles.STYLE_NORMAL_BOLD, DefaultCellStyles.normalBoldCellStyle(workbook));
        tableDefinition
                .customCellStyles(workbook)
                .forEach((key, style) -> registerCellStyle(globalConsumer, key, style));
    }

    private void registerCellStyle(Consumer<CellStyle> globalConsumer, String key, CellStyle cellStyle) {
        if (globalConsumer != null) {
            globalConsumer.accept(cellStyle);
        }
        cellStyleMap.put(key, cellStyle);
    }

    public CellStyle getCellStyle(String key) {
        return cellStyleMap.getOrDefault(key, cellStyleMap.get(DefaultCellStyles.STYLE_NORMAL));
    }
}
