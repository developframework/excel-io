package com.github.developframework.excel;

import lombok.Getter;

/**
 * @author qiushui on 2019-05-18.
 */
@Getter
public class TableLocation {

    private final int row;

    private final int column;

    private TableLocation(int row, int column) {
        this.row = row;
        this.column = column;
    }

    public static TableLocation of(int row, int column) {
        return new TableLocation(row, column);
    }
}
