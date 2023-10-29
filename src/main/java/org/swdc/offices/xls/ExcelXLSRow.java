package org.swdc.offices.xls;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.swdc.offices.CellPresetFunction;
import org.swdc.offices.RowIteratorFunction;
import org.swdc.offices.xlsx.ExcelCell;
import org.swdc.offices.xlsx.ExcelRow;

import java.util.Collection;

public class ExcelXLSRow {

    private HSSFRow row;

    private ExcelXLSSheet sheet;

    private CellPresetFunction preset;

    public ExcelXLSRow(ExcelXLSSheet sheet, HSSFRow row) {
        this.sheet = sheet;
        this.row = row;
    }

    public ExcelXLSRow presetCell(CellPresetFunction preset) {
        this.preset = preset;
        return this;
    }

    public ExcelXLSRow rowAt(int idx) {
        if (idx < 0) {
            throw new RuntimeException("invalid row number : " + idx);
        }
        HSSFSheet sheet = this.row.getSheet();
        HSSFRow xsRow = sheet.getRow(idx);
        if (xsRow == null) {
            xsRow = sheet.createRow(idx);
        }
        return new ExcelXLSRow(this.sheet,xsRow);
    }

    public ExcelXLSRow prevRow(int num) {
        return rowAt(this.row.getRowNum() - num);
    }

    public ExcelXLSRow prevRow() {
        return rowAt(this.row.getRowNum() - 1);
    }

    public ExcelXLSRow nextRow(int num) {
        return rowAt(this.row.getRowNum() + num);
    }

    public ExcelXLSRow nextRow() {
        return rowAt(this.row.getRowNum() + 1);
    }

    public ExcelXLSRow height(int num) {
        this.row.setHeight((short) num);
        return this;
    }

    public <E> ExcelXLSRow forOf(Collection<E> collection, RowIteratorFunction<E,ExcelXLSCell> func) {
        ExcelXLSRow cur = this;
        for (E e: collection) {
            func.accept(cur.cell(0),e);
            cur = cur.nextRow();
        }
        return cur;
    }

    public ExcelXLSCell cell(int col) {
        HSSFCell cell = row.getCell(col);
        if (cell == null) {
            cell = row.createCell(col);
        }
        ExcelXLSCell exCell = new ExcelXLSCell(this,cell);
        if (this.preset != null) {
            return exCell.preset(preset);
        }
        return exCell;
    }


}
