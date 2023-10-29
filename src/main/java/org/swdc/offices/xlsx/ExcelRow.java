package org.swdc.offices.xlsx;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.swdc.offices.CellPresetFunction;
import org.swdc.offices.RowIteratorFunction;

import java.util.Collection;

public class ExcelRow {

    private XSSFRow row;

    private CellPresetFunction preset;

    private ExcelSheet sheet;

    public ExcelRow(ExcelSheet sheet,XSSFRow row) {
        this.row = row;
    }

    public ExcelRow presetCell(CellPresetFunction preset) {
        this.preset = preset;
        return this;
    }

    public ExcelRow rowAt(int idx) {
        if (idx < 0) {
            throw new RuntimeException("invalid row number : " + idx);
        }
        XSSFSheet sheet = this.row.getSheet();
        XSSFRow xsRow = sheet.getRow(idx);
        if (xsRow == null) {
            xsRow = sheet.createRow(idx);
        }
        return new ExcelRow(this.sheet,xsRow);
    }

    public ExcelRow prevRow(int num) {
        return rowAt(this.row.getRowNum() - num);
    }

    public ExcelRow prevRow() {
        return rowAt(this.row.getRowNum() - 1);
    }

    public ExcelRow nextRow(int num) {
        return rowAt(this.row.getRowNum() + num);
    }

    public ExcelRow nextRow() {
        return rowAt(this.row.getRowNum() + 1);
    }

    public ExcelRow height(int num) {
        this.row.setHeight((short) num);
        return this;
    }

    public <E> ExcelRow forOf(Collection<E> collection, RowIteratorFunction<E> func) {
        ExcelRow cur = this;
        for (E e: collection) {
            func.accept(cur.cell(0),e);
            cur = cur.nextRow();
        }
        return cur;
    }

    public ExcelCell cell(int col) {
        XSSFCell cell = row.getCell(col);
        if (cell == null) {
            cell = row.createCell(col);
        }
        ExcelCell exCell = new ExcelCell(this,cell);
        if (this.preset != null) {
            return exCell.preset(preset);
        }
        return exCell;
    }

    public ExcelSheet backToSheet() {
        return sheet;
    }

}
