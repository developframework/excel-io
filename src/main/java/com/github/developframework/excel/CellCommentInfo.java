package com.github.developframework.excel;

import lombok.Getter;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;

import java.util.function.BiFunction;

/**
 * 单元格批注信息
 *
 * @author qiushui on 2023-05-19.
 */
@Getter
@RequiredArgsConstructor
public class CellCommentInfo {

    private final String authorField;

    private final String commentField;

    private final BiFunction<Drawing, Cell, ClientAnchor> anchorFunction;
}
