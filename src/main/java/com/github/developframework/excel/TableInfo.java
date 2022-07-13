package com.github.developframework.excel;

/**
 * @author qiushui on 2022-06-29.
 */
public class TableInfo {

    // 标题
    public String title;

    // 是否有标题
    public boolean hasTitle;

    // 是否有列头
    public boolean hasColumnHeader = true;

    // 工作表名称
    public String sheetName;

    // 工作表
    public int sheet;

    // 表格位置
    public TableLocation tableLocation = TableLocation.of(0, 0);
}
