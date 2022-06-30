package test;

import com.github.developframework.excel.ExcelIO;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;

/**
 * @author qiushui on 2022-06-29.
 */
public class Read {

    public static void main(String[] args) {
        final List<Student> students = ExcelIO.reader("D:\\学生成绩表.xlsx")
                .read(Student.class, (workbook, builder) ->
                        builder.columnDefinitions(
                                builder.<Student, String>column("name"),
                                builder.<Student, Student.Gender>column("gender"),
                                builder.<Student, LocalDate>column("birthday"),
                                builder.<Student, LocalDateTime>column("createTime"),
                                builder.<Student, LocalDate>column("chineseScore"),
                                builder.<Student, LocalDate>column("mathScore"),
                                builder.<Student, LocalDate>column("englishScore"),
                                builder.<Student, Integer>formula("totalScore"),
                                builder.<Student, Boolean>formula("qualified")
                                        .readConvert((student, qualified) -> qualified.equals("合格"))
                        )
                );
        students.forEach(System.out::println);
    }
}
