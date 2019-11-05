package com.github.developframework.excel.annotations;

import java.lang.annotation.*;

/**
 * Excel表
 *
 * @author qiushui on 2019-05-18.
 */
@Inherited
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelTable {

    /* 列名 */
    String title() default "";

    int row();

    int column();

}
