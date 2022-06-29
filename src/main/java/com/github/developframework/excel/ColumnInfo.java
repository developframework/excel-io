package com.github.developframework.excel;

import lombok.RequiredArgsConstructor;

/**
 * @author qiushui on 2022-06-28.
 */
@RequiredArgsConstructor
public class ColumnInfo {
    public final String field;
    public final String header;
    public Integer columnWidth;


    public ColumnInfo(String field, String header, Integer columnWidth) {
        this.field = field;
        this.header = header;
        this.columnWidth = columnWidth;
    }
}
