package com.github.developframework.excel.column;

import com.github.developframework.excel.AbstractColumnDefinition;
import org.apache.poi.ss.usermodel.Cell;

/**
 * @author qiushui on 2019-09-02.
 */
public class BlankColumnDefinition<ENTITY> extends AbstractColumnDefinition<ENTITY, Void> {

    public BlankColumnDefinition(String header) {
        super(null, header);
    }

    @Override
    protected String getCellValue(Cell cell) {
        return null;
    }
}
