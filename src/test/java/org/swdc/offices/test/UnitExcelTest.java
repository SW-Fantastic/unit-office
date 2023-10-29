package org.swdc.offices.test;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.swdc.offices.xlsx.CellPresetFunction;
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

        CellPresetFunction presetField = c -> c.font()
                .bold(true)
                .back()
                .alignVerticalCenter();

        List<String> items = Arrays.asList(
                "Arch QP", "Arch",
                "C & S QP", "EE",
                "ME", "CM", "RE",
                "RTO", "Others"
        );

        sheet.rowAt(0).presetCell(presetField)
                .cell(1).text("To:").nextCell().text("Surbana Juyron Consulatans Ple Ltd")
                .backToRow().nextRow().cell(2).text("168 Jalan Bukit merah")
                .backToRow().nextRow().cell(2).text("singepore 00000")
                .backToRow().nextRow(2).cell(1).preset(presetField).text("Attn:").nextCell().text("Mr xxx, Mr xxxx")
                .backToRow().nextRow(2).cell(1).preset(presetField).text("Thru:").nextCell().text("Ms Lee Eng mui(Manager)")
                .backToRow().prevRow(6).cell(12).preset(presetField).text("RFI Ref").nextCell().text("BHC")
                .backToRow().nextRow().cell(12).text("Date:").nextCell().text("5/23/2001")
                .backToRow().nextRow().cell(12).text("Your Fax No:").nextCell().text("")
                .backToRow().nextRow().cell(12).text("By:").nextCell().text("Fax / Email / Hard")
                .backToRow().nextRow().cell(12).text("Our fax no:").nextCell().text("")
                .backToRow().nextRow().cell(12).text("No of pages:").nextCell().text("12")
                .backToRow().prevRow(5).forOf(items, (c,e) -> c
                        .cellAt(19)
                        .border(BorderStyle.THIN)
                        .borderColor("#000")
                        .text(e)
                ).nextRow()
                .cell(0).picture().file(
                        new File("test.png"),
                        Workbook.PICTURE_TYPE_PNG
                ).cross(2,2)
                .back().backToRow();

        workbook.write(new FileOutputStream("test.xlsx"));
    }

}
