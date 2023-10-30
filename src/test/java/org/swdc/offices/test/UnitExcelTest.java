package org.swdc.offices.test;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.swdc.offices.CellPresetFunction;
import org.swdc.offices.xlsx.ExcelCell;
import org.swdc.offices.xlsx.ExcelSheet;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

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

    public static void main(String[] args) throws IOException {

        XSSFWorkbook workbook = new XSSFWorkbook();
        ExcelSheet sheet = new ExcelSheet(workbook,"Sheet A");

        CellPresetFunction<ExcelCell> presetField = c -> c.font()
                .bold(true)
                .back()
                .alignVerticalCenter();

        List<String> items = Arrays.asList(
                "Arch QP", "Arch",
                "C & S QP", "EE",
                "ME", "CM", "RE",
                "RTO", "Others"
        );


        for (int j = 25+15;j<24+15+29;j++){
            for (int i = 3;i<=18;i++){
                sheet.rowAt(j).cell(i).borderTop(BorderStyle.MEDIUM).borderColor("#000").backToRow()
                        .cell(i).borderLeft(BorderStyle.MEDIUM).borderColor("#000").backToRow()
                        .cell(i).borderRight(BorderStyle.MEDIUM).borderColor("#000").backToRow()
                        .cell(i).borderBottom(BorderStyle.MEDIUM).borderColor("#000");

            }
        }

        workbook.write(new FileOutputStream("test.xlsx"));
    }

}
