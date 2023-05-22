package test;

import com.github.developframework.excel.styles.CellStyleKey;

/**
 * @author qiushui on 2023-05-19.
 */
public class Main {

    public static void main(String[] args) {
        String key = "align {h: right}";
        final CellStyleKey cellStyleKey = CellStyleKey.parse(key);
        System.out.println(cellStyleKey);
    }
}
