# UnitOffice

本项目是POI的一个Wrapper，目的是更方便的使用POI操作Excel文档，
本项目提供了非常流畅的API，能够自如的应对各种Excel的生成操作，
目前项目还在积极的开发中，不建议直接使用。

本项目采用MIT协议提供。

# Quick start

通常，如果你需要使用POI生成Excel，那么这将会是一个很繁琐的过程，
你需要创建WorkBook，然后分别处理对应的行列以及Cell的样式，这是
一个十分费时费力的过程，本项目的目的就是将这一过程变得简单而流畅。

下面，我将演示如何通过本项目进行基本的Excel处理：

首先，这里有一个非常简单的POJO类型——Person
```java
public class Person {

    private String name;

    private String age;

    private String gender;

    private String birthDay;

    // 省略Getter Setter
    
}
```

那么对于这样的对象，通过本项目进行导出，是非常简单的：

```java
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.swdc.offices.xlsx.ExcelRow;
import org.swdc.offices.xlsx.ExcelSheet;

import java.io.FileOutputStream;

public class Demo {

    public static void main(String[] args) {

        // 第一步，创建一个Workbook
        XSSFWorkbook workbook = new XSSFWorkbook();
        // 第二步，创建ExcelSheet对象，如果你在使用HSSF类型，这里需要的
        // 就是ExcelXLSSheet，但是没有关系，它们的API是一致的。
        ExcelSheet sheet = new ExcelSheet(workbook, "Sheet A");
        // 第三步，创建表头，
        ExcelRow row = sheet.rowAt(0) // 取第零行，第零个cell，开始填写内容
                .cell(0).text("姓名")
                .nextCell().text("年龄")
                .nextCell().text("生日")
                .nextCell().text("性别")
                .backToRow(); // 返回Row，接下来可以继续对行进行操作。


        // 这里有一组数据，它们每一个都可以生成一行内容
        List<Person> personList = Arrays.asList(
                new Person("张三", "20", "Male", "2021/3/1"),
                new Person("张三3", "20", "Male", "2021/3/3"),
                new Person("张三1", "20", "Male", "2021/4/1"),
                new Person("张三4", "20", "Male", "2021/6/1"),
                new Person("张三6", "20", "Male", "2021/3/8"),
                new Person("张三5", "20", "Male", "2021/7/6")
        );
        row = row.nextRow() // 跳转到表头的下一行，使用ForOf循环生成内容
                .forOf(personList, (person, cell) -> {
                    // 回调会提供当前生成的对象以及本行第0个Cell
                    // 这时要做的很简单，在对应Cell填写数据即可。
                    cell.text(person.getName()).nextCell()
                            .text(person.getAge()).nextCell()
                            .text(person.getBirthDay()).nextCell()
                            .text(person.getGender());
                });
        // 到此就生成完毕了，接下来只需要写到必要的位置，Excel的生成流程也就结束了。
        workbook.write(new FileOutputStream("testPerson.xlsx"));
    }

}
```
那么如果遇到很复杂的Excel，又该怎么办呢？本项目也提供了对应的设计，虽然生成Excel的
逻辑比较繁琐，但是如果以合适的形式进行组织，那么它的可维护性是非常有保证的：

```java
import org.swdc.offices.generator.GeneratorStage;

public class DemoStrategies {

    @GeneratorStage
    public ExcelRow generateHeader(PipedGenerationContext ctx, ExcelSheet sheet) {
        sheet.autoColumnWidth(0)
                .autoColumnWidth(1)
                .autoColumnWidth(2)
                .autoColumnWidth(3);

        CellPresetFunction<ExcelCell> preset = cell -> cell
                .font()
                .bold(true)
                .back()
                .alignCenter();

        return sheet.rowAt(0).presetCell(preset)
                .cell(0).text("姓名")
                .nextCell().text("年龄")
                .nextCell().text("生日")
                .nextCell().text("性别")
                .backToRow();
    }

    @GeneratorStage
    public void generatePerson(PipedGenerationContext ctx, ExcelSheet sheet) {
        sheet.rowAt(1).forOf(ctx.getGrouped(Person.class), (cell, person) -> {
            cell.text(person.getName()).nextCell()
                    .text(person.getAge()).nextCell()
                    .text(person.getBirthDay()).nextCell()
                    .text(person.getGender());
        });
    }

}
```
如上所示，不仅可以生成Excel，而且API还可以对Excel的格式进行调整，当然这部分功能目前还相对有限，
通过GeneratorStage注解，一个复杂Excel的生成可以以一定的标准分为不同的部分，每一个方法承担
生成其中一部分的功能，接下来只需要通过PipedExcelGenerator就能使用它创建Excel了：

```java
import org.swdc.offices.generator.PipedExcelGenerator;

import java.io.FileOutputStream;

public class Demo {

    public static void main(String[] args) {
        // 这里有一组数据，它们每一个都可以生成一行内容
        List<Person> personList = Arrays.asList(
                new Person("张三", "20", "Male", "2021/3/1"),
                new Person("张三3", "20", "Male", "2021/3/3"),
                new Person("张三1", "20", "Male", "2021/4/1"),
                new Person("张三4", "20", "Male", "2021/6/1"),
                new Person("张三6", "20", "Male", "2021/3/8"),
                new Person("张三5", "20", "Male", "2021/7/6")
        );

        PipedExcelGenerator generator = new PipedExcelGenerator();
        // 直接使用之前DemoStrategies对象
        generator.generateStages(new DemoStrategies());
        // 通过本方法创建Excel。
        generator.createExcel("Sheet A", personList, new FileOutputStream("demoExcel2.xlsx"));
    }

}
```