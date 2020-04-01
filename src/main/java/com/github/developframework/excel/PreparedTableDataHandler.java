package com.github.developframework.excel;

import java.util.List;

/**
 * 表数据预处理器
 *
 * @param <S>
 * @param <T>
 */
public interface PreparedTableDataHandler<S, T> {

    List<T> handle(List<S> source);
}
