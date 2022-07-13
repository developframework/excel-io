package com.github.developframework.excel.column;

import com.github.developframework.excel.AbstractColumnDefinition;

/**
 * 通用列
 *
 * @author qiushui on 2019-05-18.
 */
public class GeneralColumnDefinition<ENTITY, FIELD> extends AbstractColumnDefinition<ENTITY, FIELD> {

    protected GeneralColumnDefinition(String field, String header) {
        super(field, header);
    }
}
