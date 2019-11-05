/**
 * @author qiushui on 2019-05-19.
 */
public class Test {

//    public static void main(String[] args) {
//        TableDefinition tableDefinition = (workbook, builder) -> builder.columnDefinitions(
//                builder.string("name", "姓名"),
//                builder.numeric("age", "年龄")
//                        .valueToDouble(int.class, v -> (double) v * 2)
//                        .doubleToValue(int.class, v -> (int) (v / 2) ),
//                builder.blank("空列"),
//                builder.string("createTime", "创建时间")
//                        .valueToString(LocalDateTime.class, DateTimeAdvice::format)
//                        .stringToValue(LocalDateTime.class, DateTimeAdvice::parseStandard),
//                builder.formula("compute", "计算").formula("B{row}*10 + {column}")
//        );
//
////        List<User> users = List.of(
////                new User("张三", 20, LocalDateTime.now()),
////                new User("李四", 25, LocalDateTime.now())
////        );
////        ExcelIO
////                .writer(ExcelType.XLSX)
////                .load(users, tableDefinition)
////                .writeToFile("E:\\test.xlsx");
//
//        List<User> users = ExcelIO
//                .reader("E:\\test.xlsx")
//                .read(User.class, tableDefinition);
//
//        users.forEach(System.out::println);
//    }
}
