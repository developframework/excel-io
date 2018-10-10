import com.github.developframework.excel.AbstractTableDefinition;
import com.github.developframework.excel.ExcelIO;
import com.github.developframework.excel.ExcelType;
import com.github.developframework.excel.TableDefinition;
import com.github.developframework.excel.column.BasicColumnDefinition;
import com.github.developframework.excel.column.ColumnDefinition;
import com.github.developframework.excel.column.DateTimeColumnDefinition;
import com.github.developframework.excel.column.NumberColumnDefinition;
import com.github.developframework.mock.MockClient;
import org.apache.commons.lang3.time.DateUtils;
import org.apache.poi.ss.usermodel.Workbook;

import java.text.ParseException;
import java.util.List;

/**
 * @author qiushui on 2018-10-10.
 * @since 0.1
 */
public class Main {

    public static void main(String[] args) {

        MockClient mockClient = new MockClient();

        Customer[] customers = new Customer[100];
        for (int i = 0; i < customers.length; i++) {
            customers[i] = new Customer();

            customers[i].setName(mockClient.mock("${ personName | length=3 }"));
            customers[i].setMoney(100);
            customers[i].setMobile(mockClient.mock("${ mobile }"));
            try {
                customers[i].setBirthday(DateUtils.parseDate(mockClient.mock("${ date | range=20y, pattern=yyyy-MM-dd }"), "yyyy-MM-dd"));
            } catch (ParseException e) {
                e.printStackTrace();
            }
        }

        TableDefinition tableDefinition = new AbstractTableDefinition() {

            @Override
            public int row() {
                return 2;
            }

            @Override
            public int column() {
                return 1;
            }

            @Override
            public ColumnDefinition[] columnDefinitions(Workbook workbook) {

                return new ColumnDefinition[]{
                        new BasicColumnDefinition(workbook, "姓名", "name"),
                        new BasicColumnDefinition(workbook, "手机号", "mobile"),
                        new DateTimeColumnDefinition(workbook, "出生日期", "birthday", "yyyy-MM-dd"),
                        new NumberColumnDefinition(workbook, "金额", "money", "￥0.00"),
                };
            }
        };

        ExcelIO.writer(ExcelType.XLSX, "E:\\test.xlsx")
                .fillData(customers, tableDefinition)
                .write();
    }

    public static void main1(String[] args) {
        List<Customer> list = ExcelIO.reader(ExcelType.XLSX, "E:\\test.xlsx")
                .readAndClose(Customer.class, null, new AbstractTableDefinition() {

                    @Override
                    public int row() {
                        return 5;
                    }

                    @Override
                    public int column() {
                        return 4;
                    }

                    @Override
                    public Integer sheet() {
                        return 0;
                    }

                    @Override
                    public ColumnDefinition[] columnDefinitions(Workbook workbook) {

                        return new ColumnDefinition[]{
                                new BasicColumnDefinition(workbook, "姓名", "name"),
                                new BasicColumnDefinition(workbook, "手机号", "mobile"),
                                new DateTimeColumnDefinition(workbook, "出生日期", "birthday", "yyyy-MM-dd"),
                                new NumberColumnDefinition(workbook, "金额", "money", "￥0.00"),
                        };
                    }
                });

        list.forEach(System.out::println);
    }
}
