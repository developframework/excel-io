package test;

import com.github.developframework.excel.styles.CellStyleKey;

/**
 * @author qiushui on 2023-05-19.
 */
public class Main {

    public static void main(String[] args) {
        String key = "f {size: 16;bold; family: 宋体; color: #ffaaee} a {v:bottom} fg{color: #aa1199;} b{style: thin;color: #bbaadd;} df{format: 0.00%}";
        System.out.println(CellStyleKey.isCellStyleKey(key));


        final CellStyleKey cellStyleKey = CellStyleKey.parse(key);
        System.out.println(cellStyleKey);
    }
}
