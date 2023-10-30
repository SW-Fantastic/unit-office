package org.swdc.offices.xls;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;
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

    public ExcelXLSSheet autoColumnWidth(int column) {
        if (column < 0) {
            throw new RuntimeException("column index is incorrect.");
        }
        sheet.autoSizeColumn(column);
        return this;
    }

    public ExcelXLSSheet columnVisible(int column,boolean visible) {
        if (column < 0) {
            throw new RuntimeException("column index is invalid");
        }
        sheet.setColumnHidden(column,visible);
        return this;
    }

    public ExcelXLSSheet mergeCells(int row, int column, int rowSpan, int colSpan) {
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

    public ExcelXLSSheet splitMergeCells(int row, int column, int rowSpan, int columnSpan) {
        for( CellRangeAddress addr : sheet.getMergedRegions() ) {
            if (
                    addr.getFirstRow() == row &&
                            addr.getFirstColumn() == column &&
                            addr.getLastRow() == row + rowSpan &&
                            addr.getLastColumn() == column + columnSpan
            ) {
                sheet.removeMergedRegion(sheet.getMergedRegions().indexOf(addr));
                return this;
            }
        }
        return this;
    }

}
