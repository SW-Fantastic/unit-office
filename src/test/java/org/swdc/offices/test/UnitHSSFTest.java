package org.swdc.offices.test;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.swdc.offices.CellPresetFunction;
import org.swdc.offices.xls.ExcelXLSCell;
import org.swdc.offices.xls.ExcelXLSSheet;
import org.swdc.offices.xlsx.ExcelCell;
import org.swdc.offices.xlsx.ExcelSheet;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

public class UnitHSSFTest {

    public static void main(String[] args) throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        ExcelXLSSheet sheet = new ExcelXLSSheet(workbook,"Sheet A");

        CellPresetFunction<ExcelXLSCell> presetField = c -> c.font()
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
                        new File("test.png")
                ).cross(2,2)
                .back().backToRow();

        workbook.write(new FileOutputStream("test.xls"));
    }

}
