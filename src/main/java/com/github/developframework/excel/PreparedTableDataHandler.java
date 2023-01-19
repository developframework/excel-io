package com.github.developframework.excel;

import java.util.List;

/**
 * 表数据预处理器
 *
 * @param <S>
 */
public interface PreparedTableDataHandler<S> {

    void handle(List<S> source);
}
