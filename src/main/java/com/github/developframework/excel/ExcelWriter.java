package com.github.developframework.excel;

import com.github.developframework.excel.column.ColumnDefinitionBuilder;
import com.github.developframework.excel.styles.CellStyleManager;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;

import java.io.*;
import java.util.Arrays;
import java.util.List;

/**
 * @author qiushui on 2019-05-18.
 */
@SuppressWarnings("unused")
public class ExcelWriter extends ExcelProcessor {

    protected ExcelWriter(Workbook workbook) {
        super(workbook);
    }

    /**
     * 填充数据
     *
     * @param data            实体列表
     * @param tableDefinition 表定义
     * @return 写出器
     */
    public <ENTITY> ExcelWriter load(List<ENTITY> data, TableDefinition<ENTITY> tableDefinition) {
        writeInternal(tableDefinition, data);
        return this;
    }

    /**
     * 填充数据
     *
     * @param data            实体列表
     * @param tableDefinition 表定义
     * @return 写出器
     */
    public <ENTITY> ExcelWriter load(ENTITY[] data, TableDefinition<ENTITY> tableDefinition) {
        writeInternal(tableDefinition, Arrays.asList(data));
        return this;
    }

    /**
     * 写出
     */
    public void write(OutputStream outputStream) {
        try {
            workbook.write(outputStream);
            workbook.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 写出到文件
     *
     * @param filename 文件名
     */
    public File writeToFile(String filename) {
        File file = new File(filename);
        try (OutputStream os = new FileOutputStream(file)) {
            write(os);
            return file;
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 写出到字节数组
     *
     * @return 字节数组
     */
    public byte[] writeToByteArray() {
        try (ByteArrayOutputStream os = new ByteArrayOutputStream()) {
            write(os);
            return os.toByteArray();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 写入表格
     *
     * @param tableDefinition 标定仪
     * @param list            实体列表
     * @param <ENTITY>        实体类泛型
     */
    @SuppressWarnings({"unchecked", "rawtypes"})
    private <ENTITY> void writeInternal(TableDefinition<ENTITY> tableDefinition, List<ENTITY> list) {
        final PreparedTableDataHandler<ENTITY> preparedTableDataHandler = (PreparedTableDataHandler<ENTITY>) tableDefinition.preparedTableDataHandler();
        if (preparedTableDataHandler != null) {
            preparedTableDataHandler.handle(list);
        }

        final TableInfo tableInfo = tableDefinition.tableInfo();

        // 创建工作表
        final Sheet sheet = createSheet(tableInfo);

        // 列定义
        final ColumnDefinition<ENTITY>[] columnDefinitions = tableDefinition.columnDefinitions(workbook, new ColumnDefinitionBuilder(workbook));

        // 表位置
        final int startColumnIndex = tableInfo.tableLocation.getColumn();
        int rowIndex = tableInfo.tableLocation.getRow();

        // 单元格样式管理器
        final CellStyleManager cellStyleManager = new CellStyleManager(workbook, tableDefinition);

        // 创建表头标题
        if (tableInfo.hasTitle && tableInfo.title != null) {
            tableDefinition.createTableTitle(sheet, cellStyleManager, rowIndex++, startColumnIndex, tableInfo.title, columnDefinitions.length);
        }
        // 创建列头
        if (tableInfo.hasColumnHeader) {
            rowIndex = tableDefinition.createTableColumnHeader(sheet, cellStyleManager, rowIndex, startColumnIndex, columnDefinitions);
        }

        // 创建表主体数据
        tableDefinition.createTableBody(workbook, sheet, cellStyleManager, rowIndex, startColumnIndex, columnDefinitions, list);

        final int maxWidth = 255 * 256;
        // 设置列宽
        for (int i = 0; i < columnDefinitions.length; i++) {
            int col = startColumnIndex + i;
            final ColumnInfo columnInfo = columnDefinitions[i].getColumnInfo();
            if (columnInfo != null) {
                if (columnInfo.columnWidth == null) {
                    // 自动列宽
                    sheet.autoSizeColumn(col);
                    // 解决自动设置列宽中文失效的问题
                    sheet.setColumnWidth(col, Math.min(maxWidth, sheet.getColumnWidth(i)));
                } else {
                    sheet.setColumnWidth(col, Math.min(maxWidth, columnInfo.columnWidth * 256));
                }
            }
        }

        // 工作表额外处理
        SheetExtraHandler sheetExtraHandler = tableDefinition.sheetExtraHandler();
        if (sheetExtraHandler != null) {
            sheetExtraHandler.handle(workbook, sheet, cellStyleManager, rowIndex, rowIndex + list.size(), list);
        }
    }

    /**
     * 创建工作表
     *
     * @param tableInfo 表格信息
     * @return 工作表
     */
    private Sheet createSheet(TableInfo tableInfo) {
        Sheet sheet;
        if (tableInfo.sheetName == null) {
            sheet = workbook.createSheet();
        } else {
            sheet = workbook.createSheet(tableInfo.sheetName);
        }
        if (sheet instanceof SXSSFSheet) {
            ((SXSSFSheet) sheet).setRandomAccessWindowSize(-1);
            // 开启追踪列宽
            ((SXSSFSheet) sheet).trackAllColumnsForAutoSizing();
        }
        return sheet;
    }
}
