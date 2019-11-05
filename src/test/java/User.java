import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.time.LocalDateTime;

/**
 * @author qiushui on 2019-05-19.
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
public class User {

    private String name;

    private int age;

    private LocalDateTime createTime;

    private int compute;
}
