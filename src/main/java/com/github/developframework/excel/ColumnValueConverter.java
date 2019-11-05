package com.github.developframework.excel;

/**
 * @author qiushui on 2019-05-18.
 */
public interface ColumnValueConverter<ENTITY, SOURCE, TARGET> {

    TARGET convert(ENTITY entity, SOURCE value);
}
