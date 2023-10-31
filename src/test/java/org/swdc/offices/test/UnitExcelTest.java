package org.swdc.offices.test;

import org.swdc.offices.CellPresetFunction;
import org.swdc.offices.generator.PipedExcelGenerator;
import org.swdc.offices.generator.PipedGenerationContext;
import org.swdc.offices.xlsx.ExcelCell;
import org.swdc.offices.generator.ExcelSimpleGenerator;
import org.swdc.offices.xlsx.ExcelRow;
import org.swdc.offices.xlsx.ExcelSheet;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;

public class UnitExcelTest {

    public static class Person {

        private String name;

        private String age;

        private String gender;

        private String birthDay;

        public Person(String name, String age, String gender, String birthDay) {
            this.age = age;
            this.birthDay = birthDay;
            this.gender = gender;
            this.name = name;
        }

        public String getName() {
            return name;
        }

        public String getAge() {
            return age;
        }

        public String getGender() {
            return gender;
        }

        public String getBirthDay() {
            return birthDay;
        }

    }


    public static class DemoGenerator {

        public ExcelRow generateHeader(ExcelSheet sheet) {
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

        public void generatePerson(ExcelCell cell, Person person) {
            cell.text(person.getName()).nextCell()
                    .text(person.getAge()).nextCell()
                    .text(person.getBirthDay()).nextCell()
                    .text(person.getGender());
        }

        public ExcelSimpleGenerator createGenerator() {
            return new ExcelSimpleGenerator()
                    .generateExcelStructure(this::generateHeader)
                    .strategy(Person.class, this::generatePerson);
        }

    }


    public static class PipedDemoGenerator {

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

        public void generatePerson(PipedGenerationContext ctx, ExcelSheet sheet) {
            sheet.rowAt(1).forOf(ctx.getGrouped(Person.class), (cell, person) -> {
                cell.text(person.getName()).nextCell()
                        .text(person.getAge()).nextCell()
                        .text(person.getBirthDay()).nextCell()
                        .text(person.getGender());
            });
        }

        public PipedExcelGenerator createGenerator() {
            return new PipedExcelGenerator()
                    .generateStage(this::generateHeader)
                    .generateStage(this::generatePerson);
        }

    }

    public static void main(String[] args) throws IOException {

        List<Person> personList = Arrays.asList(
                new Person("张三","20","Male","2021/3/1"),
                new Person("张三3","20","Male","2021/3/3"),
                new Person("张三1","20","Male","2021/4/1"),
                new Person("张三4","20","Male","2021/6/1"),
                new Person("张三6","20","Male","2021/3/8"),
                new Person("张三5","20","Male","2021/7/6")
        );

        DemoGenerator generator = new DemoGenerator();
        generator.createGenerator()
                .createExcel(
                        "SheetA",
                        personList,
                        new FileOutputStream("test2.xlsx")
                );



        PipedExcelGenerator pipedDemoGenerator = new PipedDemoGenerator().createGenerator();
        pipedDemoGenerator.createExcel("Sheet A",personList,new FileOutputStream("test3.xlsx"));
    }

}
