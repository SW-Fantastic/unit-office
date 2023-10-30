package org.swdc.offices.xlsx;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelSheet {

    private XSSFSheet sheet;

    public ExcelSheet(XSSFWorkbook workbook, String name) {
        XSSFSheet sheet = workbook.getSheet(name);
        if (sheet == null) {
            sheet = workbook.createSheet(name);
        }
        this.sheet = sheet;
    }

    public ExcelSheet(XSSFWorkbook workbook, int index) {
        XSSFSheet sheet = workbook.getSheetAt(index);
        if (sheet == null) {
            throw new RuntimeException("no such sheet index : " + index);
        }
        this.sheet = sheet;
    }

    public ExcelRow rowAt(int row) {
        if (row < 0) {
            throw new RuntimeException("invalid row number: " + row);
        }
        XSSFRow xsRow = sheet.getRow(row);
        if (xsRow == null) {
            xsRow = sheet.createRow(row);
        }
        return new ExcelRow(this,xsRow);
    }

    public ExcelSheet columnWidth(int column, int val) {
        if (val < 0) {
            throw new RuntimeException("column can not less than zero!");
        }
        sheet.setColumnWidth(column,val);
        return this;
    }

    public ExcelSheet mergeCells(int row,int column, int rowSpan, int colSpan) {
        CellRangeAddress address = new CellRangeAddress(row,row + rowSpan,column,column + colSpan);
        for( CellRangeAddress addr : sheet.getMergedRegions() ) {
            if (
                    addr.getFirstRow() == address.getFirstRow() &&
                    addr.getFirstColumn() == address.getFirstColumn() &&
                    addr.getLastRow() == address.getLastRow() &&
                    addr.getLastColumn() == address.getLastColumn()
            ) {
                return this;
            }
        }
        sheet.addMergedRegion(address);
        return this;
    }

}
