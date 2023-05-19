package test;

import com.github.developframework.excel.styles.CellStyleKey;

/**
 * @author qiushui on 2023-05-19.
 */
public class Main {

    public static void main(String[] args) {
        String key = "f#16-BOLD-宋体-#ffaaee; a#v:bottom; fg##aa1199; b#thin-#bbaadd; df#0.00%";
        final CellStyleKey cellStyleKey = CellStyleKey.parse(key);
        System.out.println(cellStyleKey);
    }
}
