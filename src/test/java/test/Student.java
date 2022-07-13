package test;

import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import lombok.ToString;

import java.time.LocalDate;
import java.time.LocalDateTime;

/**
 * @author qiushui on 2022-06-28.
 */
@Getter
@Setter
@NoArgsConstructor
@ToString
public class Student {

    // 姓名
    private String name;

    // 性别
    private Gender gender;

    // 生日
    private LocalDate birthday;

    // 入学时间
    private LocalDateTime createTime;

    // 语文成绩
    private int chineseScore;

    // 数学成绩
    private int mathScore;

    // 英语成绩
    private int englishScore;

    // 总成绩
    private int totalScore;

    // 是否合格
    private Boolean qualified;

    public Student(String name, Gender gender, LocalDate birthday, LocalDateTime createTime, int chineseScore, int mathScore, int englishScore) {
        this.name = name;
        this.gender = gender;
        this.birthday = birthday;
        this.createTime = createTime;
        this.chineseScore = chineseScore;
        this.mathScore = mathScore;
        this.englishScore = englishScore;
    }

    public enum Gender {

        MALE, FEMALE
    }
}
