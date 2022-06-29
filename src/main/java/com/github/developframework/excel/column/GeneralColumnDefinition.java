package com.github.developframework.excel.column;

import com.github.developframework.excel.AbstractColumnDefinition;
import org.apache.poi.ss.usermodel.Cell;

/**
 * 通用列
 *
 * @author qiushui on 2019-05-18.
 */
public class GeneralColumnDefinition<ENTITY, FIELD> extends AbstractColumnDefinition<ENTITY, FIELD> {

    protected GeneralColumnDefinition(String field, String header) {
        super(field, header);
    }

    @Override
    protected String getCellValue(Cell cell) {
        final String value = dataFormatter.formatCellValue(cell);
        return value != null ? value.trim() : null;
    }
}
