package test;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;

import java.time.LocalDate;

/**
 * @author qiushui on 2022-06-28.
 */
@Getter
@NoArgsConstructor
@AllArgsConstructor
public class User {

    private String username;

    private int age;

    private Gender gender;

    private LocalDate birthday;


    public enum Gender {

        MALE, FEMALE
    }
}
