package test;

import com.github.developframework.excel.*;
import com.github.developframework.excel.column.ColumnDefinitionBuilder;
import com.github.developframework.excel.styles.DefaultCellStyles;
import org.apache.poi.ss.usermodel.*;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;
import java.util.Map;

/**
 * @author qiushui on 2022-06-28.
 */
public class Write {

    public static void main(String[] args) {
        List<Student> students = List.of(
                new Student("小赵", Student.Gender.MALE, LocalDate.of(2002, 1, 5), LocalDateTime.now(), 97, 85, 95),
                new Student("小钱", Student.Gender.FEMALE, LocalDate.of(1999, 12, 25), LocalDateTime.now(), 92, 42, 87),
                new Student("小孙", Student.Gender.MALE, LocalDate.of(2001, 6, 8), LocalDateTime.now(), 50, 67, 45),
                new Student("小李", Student.Gender.FEMALE, LocalDate.of(2003, 8, 20), LocalDateTime.now(), 80, 90, 55)
        );

        ExcelIO
                .writer(ExcelType.XLSX)
                .load(students, new TableDefinition<>() {

                            @Override
                            public Map<String, CellStyle> customCellStyles(Workbook workbook) {
                                // 设置单元格背景色
                                final CellStyle redCellStyle = DefaultCellStyles.bodyCellStyle(workbook);
                                redCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                                redCellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
                                redCellStyle.setAlignment(HorizontalAlignment.RIGHT);
                                return Map.of("redColor", redCellStyle);
                            }

                            @Override
                            public ColumnDefinition<Student>[] columnDefinitions(Workbook workbook, ColumnDefinitionBuilder<Student> builder) {
                                // 判定分数大于60
                                final CellStyleKeyProvider<Student> scoreKeyFunction = (cell, e, v) -> ((Integer) v) >= 60 ? null : "redColor";
                                // 判定分数大于180
                                final CellStyleKeyProvider<Student> totalKeyFunction = (cell, e, v) -> ((Integer) v) >= 180 ? null : "redColor";
                                // 判定是否合格
                                final CellStyleKeyProvider<Student> passKeyFunction = (cell, e, v) -> v.equals("合格") ? null : "redColor";
                                return builder.columnDefinitions(
                                        builder.<String>column("name", "学生姓名"),
                                        builder.<Student.Gender>column("gender", "性别"),
                                        builder.<LocalDate>column("birthday", "生日"),
                                        builder.<LocalDateTime>column("createTime", "入学时间"),
                                        builder.<Integer>column("chineseScore", "语文成绩").cellStyleKey(scoreKeyFunction),
                                        builder.<Integer>column("mathScore", "数学成绩").cellStyleKey(scoreKeyFunction),
                                        builder.<Integer>column("englishScore", "英语成绩").cellStyleKey(scoreKeyFunction),
                                        builder.<Integer>formula(Integer.class, "总成绩", "SUM(E{row}:G{row})").cellStyleKey(totalKeyFunction),
                                        builder.<String>formula(String.class, "是否合格", "IF(H{row} >= 180,\"合格\",\"不合格\")").cellStyleKey(passKeyFunction)
                                );
                            }
                        }
                )
                .writeToFile("D:\\学生成绩表.xlsx");

//        ExcelIO
//                .writer(ExcelType.XLSX)
//                .load(students, new TableDefinition<>() {
//
//                    /**
//                     * 设置表格信息
//                     */
//                    @Override
//                    public TableInfo tableInfo() {
//                        return new TableInfo();
//                    }
//
//                    /**
//                     * 列定义
//                     */
//                    @Override
//                    public ColumnDefinition<Student>[] columnDefinitions(Workbook workbook, ColumnDefinitionBuilder builder) {
//                        return builder.columnDefinitions(
//
//                        );
//                    }
//
//                    /**
//                     * 全局单元格样式处理
//                     */
//                    @Override
//                    public BiConsumer<Workbook, CellStyle> globalCellStylesHandle() {
//                        return (workbook, cellStyle) -> {
//                            final Font font = workbook.createFont();
//                            font.setBold(true);
//                            cellStyle.setFont(font);
//                        };
//                    }
//
//                    /**
//                     * 申明自定义单元格样式
//                     */
//                    @Override
//                    public Map<String, CellStyle> customCellStyles(Workbook workbook) {
////                        final CellStyle cellStyle = workbook.createCellStyle();
//                        final CellStyle cellStyle = DefaultCellStyles.normalCellStyle(workbook);
//                        cellStyle.setDataFormat(...);
//                        return Map.of("customKey", cellStyle);
//                    }
//
//                    /**
//                     * 工作表扩展处理
//                     */
//                    @Override
//                    public SheetExtraHandler<?> sheetExtraHandler() {
//                        return null;
//                    }
//
//                    /**
//                     * 装填完的实体单独处理
//                     */
//                    @Override
//                    public void each(Student student) {
//
//                    }
//                });
    }
}
