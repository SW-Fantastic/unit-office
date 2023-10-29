package org.swdc.offices.xls;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.swdc.offices.xlsx.ExcelRow;
import org.swdc.offices.xlsx.ExcelSheet;

public class ExcelXLSSheet {

    private HSSFSheet sheet;

    public ExcelXLSSheet(HSSFWorkbook workbook, String name) {
        HSSFSheet sheet = workbook.getSheet(name);
        if (sheet == null) {
            sheet = workbook.createSheet(name);
        }
        this.sheet = sheet;
    }

    public ExcelXLSRow rowAt(int row) {
        if (row < 0) {
            throw new RuntimeException("invalid row number: " + row);
        }
        HSSFRow xsRow = sheet.getRow(row);
        if (xsRow == null) {
            xsRow = sheet.createRow(row);
        }
        return new ExcelXLSRow(this,xsRow);
    }

    public ExcelXLSSheet columnWidth(int column, int val) {
        if (val < 0) {
            throw new RuntimeException("column can not less than zero!");
        }
        sheet.setColumnWidth(column,val);
        return this;
    }

}
