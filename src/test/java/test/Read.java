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
                                builder.<String>column("name"),
                                builder.<Student.Gender>column("gender"),
                                builder.<LocalDate>column("birthday"),
                                builder.<LocalDateTime>column("createTime"),
                                builder.<Integer>column("chineseScore"),
                                builder.<Integer>column("mathScore"),
                                builder.<Integer>column("englishScore"),
                                builder.<Integer>formula(Integer.class, "totalScore"),
                                builder.<Boolean>formula(String.class, "qualified")
                                        .readConvert((student, qualified) -> qualified.equals("合格"))
                        )
                );
        students.forEach(System.out::println);
    }
}
