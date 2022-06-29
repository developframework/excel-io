package test;

import com.github.developframework.excel.*;
import com.github.developframework.excel.column.ColumnDefinitionBuilder;
import com.github.developframework.excel.styles.DefaultCellStyles;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

import java.time.LocalDate;
import java.util.List;
import java.util.Map;
import java.util.function.BiFunction;

/**
 * @author qiushui on 2022-06-28.
 */
public class Write {

    public static void main(String[] args) {
        List<User> users = List.of(
                new User("张三", 20, User.Gender.MALE, LocalDate.now()),
                new User("李四", 21, User.Gender.FEMALE, LocalDate.now())
        );

        ExcelIO
                .writer(ExcelType.XLSX)
                .load(users, new TableDefinition<>() {

                            @Override
                            public Map<String, CellStyle> customCellStyles(Workbook workbook) {
                                final CellStyle cellStyle = DefaultCellStyles.normalCellStyle(workbook);
                                cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                                XSSFColor bgColor = new XSSFColor(new java.awt.Color(236, 255, 243), new DefaultIndexedColorMap());
                                ((XSSFCellStyle) cellStyle).setFillForegroundColor(bgColor);
                                return Map.of("color", cellStyle);
                            }

                            @Override
                            public ColumnDefinition<User>[] columnDefinitions(Workbook workbook, ColumnDefinitionBuilder builder) {
                                BiFunction<Cell, Object, String> cellStyleKey = (cell, value) -> cell.getRowIndex() % 2 == 0 ? "color" : CellStyleManager.STYLE_NORMAL;
                                return builder.columnDefinitions(

                                        builder.<User, String>column("username", "姓名")
                                                .cellStyleKey(cellStyleKey),

                                        builder.<User, Integer>column("age", "年龄")
                                                .cellStyleKey(cellStyleKey),

                                        builder.<User, User.Gender>column("gender", "性别")
                                                .writeConvert((user, gender) -> gender == User.Gender.MALE ? "男" : "女")
                                                .cellStyleKey(cellStyleKey),

                                        builder.<User, LocalDate>column("birthday", "生日")
                                                .cellStyleKey(cellStyleKey)
                                );
                            }
                        }
                )
                .writeToFile("D:\\测试.xlsx");
    }
}
