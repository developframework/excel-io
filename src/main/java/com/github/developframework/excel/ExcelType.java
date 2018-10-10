package com.github.developframework.excel;

import lombok.Getter;

/**
 * Excel类型
 *
 * @author qiushui on 2018-10-10.
 * @since 0.1
 */
public enum ExcelType {

    XLS(".xls"),
    XLSX(".xlsx");

    @Getter
    private String extensionName;

    ExcelType(String extensionName) {
        this.extensionName = extensionName;
    }
}
