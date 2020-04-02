# excel-io

封装poi对Office Excel的输入输出工具，简化简单的导入和导出Excel数据的操作。（暂不支持合并单元格）

```xml
<dependency>
    <groupId>com.github.developframework</groupId>
    <artifactId>excel-io</artifactId>
</dependency>
```

## 教程

假设存在实体`Customer`包装数据

```java
@Data
public class Customer {

    private String name;

    private LocalDate buyDate;

    private String[] tickets;

    private int cost;
}
```

### ExcelIO

使用`ExcelIO`得到输入输出处理器

#### ExcelWriter

```java
ExcelIO
    .writer(ExcelType.XLSX)
	.load(customers, tableDefinition)
    .write(outputStream);
```
#### ExcelReader

```java
List<Customer> customers = ExcelIO
    .reader(ExcelType.XLSX, inputStream)
    .read(Customer.class, tableDefinition);
```

### TableDefinition

该接口是表格的定义类，一个定义类代表了一个数据表

通过该接口可以设置表格的表头信息和表格的左上角单元格位置（工作表、行、列）。

| 可实现方法                                           | 说明                    | 默认值 |
| ---------------------------------------------------- | ----------------------- | ------ |
| hasTitle()                                           | 表格顶部是否含有标题    | false  |
| title()                                              | 标题文本                | null   |
| hasColumnHeader()                                    | 是否有列说明            | true   |
| sheetName()                                          | 工作表名称              | null   |
| tableLocation()                                      | 表格位置(起始行,起始列) | null   |
| columnDefinitions(Workbook, ColumnDefinitionBuilder) | 列定义                  | 未实现 |
| sheetExtraHandler()                                  | 工作表其它扩展处理      | null   |

### ColumnDefinition

该抽象类是表格的列定义类，一个定义类代表了表中的某一列，指代了一个字段

内置有如下实现类：

+ StringColumnDefinition

  文本列，默认单元格格式为加边框，文字居中，单元格类型为STRING

+ FormulaColumnDefinition

  公式列，默认单元格格式为加边框，文字居中，单元格类型为FORMULA

+ NumericColumnDefinition

  数值列，默认单元格格式为加边框，文字居中，单元格类型为NUMERIC

+ BlankColumnDefintion

  空白列，默认单元格格式为加边框，文字居中，单元格类型为BLANK

+ MultipleLinesColumnDefinition

  多行文本列，默认单元格格式为加边框，文字居中，单元格类型为STRING

## 导入数据到Excel

使用`excel-io`导入customer数据

```java
Customer[] customers = new Customer[100];
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
		// 申明列
        return new ColumnDefinition[] {
                new BasicColumnDefinition(workbook, "姓名", "name"),
                new BasicColumnDefinition(workbook, "手机号", "mobile"),
                new DateTimeColumnDefinition(workbook, "出生日期", "birthday", "yyyy-MM-dd"),
                new NumberColumnDefinition(workbook, "金额", "money", "￥0.00")
        };
    }
};

ExcelIO.writer(ExcelType.XLSX, "E:\\test.xlsx")
        .fillData(customers, tableDefinition)
        .write();
```

## 从Excel导出数据

使用`excel-io`导出customer数据

```java
List<Customer> customers = new LinkedList<>();
// 准备数据略
ExcelIO
        .writer(ExcelType.XLSX)
        .load(customers, (workbook, builder) -> builder.columnDefinitions(
                builder.string("name", "姓名"),
                builder.string("buyDate", "购买日期"),
                builder.multipleLines("tickets", "所购门票"),
                builder.numeric("cost", "花费").format("￥0.00")
        ))
        .write(outputStream);
```
![](doc-images/image1.png)
