package test;

import com.github.developframework.excel.ExcelIO;

import java.util.List;

/**
 * @author qiushui on 2022-06-29.
 */
public class Read {

    public static void main(String[] args) {
        final List<User> users = ExcelIO.reader("D:\\测试.xlsx")
                .read(User.class, (workbook, builder) ->
                                builder.columnDefinitions(
                                        builder.column("username"),
                                        builder.column("age")
//                                builder.column("birthday"),
//                                builder.column("gender")
                                )
                );
        System.out.println(users);
    }
}
