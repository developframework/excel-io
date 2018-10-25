import lombok.Data;

import java.util.Date;

/**
 * @author qiushui on 2018-10-10.
 * @since 0.1
 */
@Data
public class Customer {

    private String name;

    private Date birthday;

    private String mobile;

    private int money;
}
