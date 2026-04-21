/**
 * @author qiushui on 2023-01-19.
 */
module excel.io {
    requires expression;
    requires lombok;
    requires org.apache.commons.lang3;
    requires org.slf4j;
    requires java.desktop;
    requires org.apache.poi.poi;
    requires org.apache.poi.ooxml;

    exports com.github.developframework.excel.column;
    exports com.github.developframework.excel.styles;
    exports com.github.developframework.excel.utils;
    exports com.github.developframework.excel;

}