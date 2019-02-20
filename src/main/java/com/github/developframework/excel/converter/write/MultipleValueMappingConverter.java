package com.github.developframework.excel.converter.write;

import com.github.developframework.excel.column.ColumnDefinition;

import java.util.Collection;
import java.util.List;
import java.util.Objects;
import java.util.Set;
import java.util.function.Function;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * 多值映射转换器
 *
 * @author qiushui on 2019-01-18.
 */
public class MultipleValueMappingConverter<T, R> implements ColumnDefinition.ColumnValueConverter {

    private Function<T, R> mapper;

    public MultipleValueMappingConverter(Function<T, R> mapper) {
        this.mapper = mapper;
    }

    @Override
    @SuppressWarnings("unchecked")
    public Object convert(Object data, Object currentValue) {
        Stream<T> stream;
        if (currentValue.getClass().isArray()) {
            stream = Stream.of((T[]) currentValue);
        } else if (currentValue instanceof List | currentValue instanceof Set) {
            stream = ((Collection) currentValue).stream();
        } else {
            throw new IllegalArgumentException();
        }
        return stream.filter(Objects::nonNull).map(mapper).filter(Objects::nonNull).collect(Collectors.toList());
    }
}
