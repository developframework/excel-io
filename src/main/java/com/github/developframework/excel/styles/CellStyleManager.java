package com.github.developframework.excel.styles;

import com.github.developframework.excel.TableDefinition;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.HashMap;
import java.util.Map;
import java.util.function.BiConsumer;

/**
 * 单元格样式管理器
 *
 * @author qiushui on 2022-06-29.
 */
public class CellStyleManager {

    private final Workbook workbook;

    private final BiConsumer<Workbook, CellStyle> globalConsumer;

    private final Map<String, CellStyle> cellStyleMap = new HashMap<>();

    public CellStyleManager(Workbook workbook, TableDefinition<?> tableDefinition) {
        this.workbook = workbook;
        this.globalConsumer = tableDefinition.globalCellStylesHandle();
        tableDefinition
                .customCellStyles(workbook)
                .forEach((key, style) -> registerCellStyle(workbook, globalConsumer, key, style));
    }

    private void registerCellStyle(Workbook workbook, BiConsumer<Workbook, CellStyle> globalConsumer, String key, CellStyle cellStyle) {
        if (globalConsumer != null) {
            globalConsumer.accept(workbook, cellStyle);
        }
        cellStyleMap.put(key, cellStyle);
    }

    public CellStyle getCellStyle(String key) {
        CellStyle cellStyle = cellStyleMap.get(key);
        if(cellStyle == null) {
            if(CellStyleKey.isCellStyleKey(key)) {
                final CellStyleKey cellStyleKey = CellStyleKey.parse(key);
                cellStyle = workbook.createCellStyle();
                cellStyleKey.configureCellStyle(workbook, cellStyle);
                registerCellStyle(workbook, globalConsumer, key, cellStyle);
            } else {
                throw new IllegalArgumentException(String.format("\"%s\" is not exists or invalid", key));
            }
        }
        return cellStyle;
    }
}
