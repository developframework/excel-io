package com.github.developframework.excel.annotations;

import java.lang.annotation.*;

/**
 * Excel列
 *
 * @author qiushui on 2019-05-18.
 */
@Inherited
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelColumn {

    /* 列名 */
    String header();

    /* 列序 */
    int column();
}
