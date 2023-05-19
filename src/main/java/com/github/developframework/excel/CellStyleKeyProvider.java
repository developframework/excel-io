package com.github.developframework.excel;

import org.apache.poi.ss.usermodel.Cell;

/**
 * @author qiushui on 2023-05-18.
 */
@FunctionalInterface
public interface CellStyleKeyProvider<ENTITY> {

    String provideCellStyleKey(Cell cell, ENTITY entity, Object value);
}
