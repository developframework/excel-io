package test;

import com.github.developframework.excel.ExcelIO;
import com.github.developframework.excel.ExcelType;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;

/**
 * @author qiushui on 2022-06-28.
 */
public class Write {

    public static void main(String[] args) {
        List<Student> students = List.of(
                new Student("小赵", Student.Gender.MALE, LocalDate.of(2002, 1, 5), LocalDateTime.now(), 97, 85, 95),
                new Student("小钱", Student.Gender.FEMALE, LocalDate.of(1999, 12, 25), LocalDateTime.now(), 92, 89, 87),
                new Student("小孙", Student.Gender.MALE, LocalDate.of(2001, 6, 8), LocalDateTime.now(), 50, 40, 45),
                new Student("小李", Student.Gender.FEMALE, LocalDate.of(2003, 8, 20), LocalDateTime.now(), 80, 90, 72)
        );

        ExcelIO
                .writer(ExcelType.XLSX)
                .load(students, (workbook, builder) ->
                        builder.columnDefinitions(
                                builder.<Student, String>column("name", "学生姓名"),
                                builder.<Student, Student.Gender>column("gender", "性别"),
                                builder.<Student, LocalDate>column("birthday", "生日"),
                                builder.<Student, LocalDateTime>column("createTime", "入学时间"),
                                builder.<Student, Integer>column("chineseScore", "语文成绩"),
                                builder.<Student, Integer>column("mathScore", "数学成绩"),
                                builder.<Student, Integer>column("englishScore", "英语成绩"),
                                builder.<Student, Integer>formula("总成绩", "SUM(E{row}:G{row})"),
                                builder.<Student, Boolean>formula("是否合格", "IF(H{row} >= 180,\"合格\",\"不合格\")")
                        )
                )
                .writeToFile("D:\\学生成绩表.xlsx");
    }
}
