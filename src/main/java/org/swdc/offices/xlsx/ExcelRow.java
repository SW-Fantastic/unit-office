package org.swdc.offices.xlsx;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.swdc.offices.CellPresetFunction;
import org.swdc.offices.RowIteratorFunction;

import java.util.Collection;

/**
 * XLSX（XSSF）表中的一行。
 */
public class ExcelRow {

    private XSSFRow row;

    private CellPresetFunction preset;

    private ExcelSheet sheet;

    public ExcelRow(ExcelSheet sheet,XSSFRow row) {
        this.sheet = sheet;
        this.row = row;
    }

    /**
     * 为本行全部的Cell设置预设样式
     * @param preset 预设样式函数
     * @return 本行
     */
    public ExcelRow presetCell(CellPresetFunction preset) {
        this.preset = preset;
        return this;
    }

    /**
     * 到指定的行，如果不存在则会创建
     * @param idx 指定行的index。
     * @return 指定行
     */
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

    /**
     * 跳转到上方的第n行
     * @param num 向上跳过的行数
     * @return 指定行
     */
    public ExcelRow prevRow(int num) {
        return rowAt(this.row.getRowNum() - num);
    }

    /**
     * 上一行
     * @return 本行的上一行
     */
    public ExcelRow prevRow() {
        return rowAt(this.row.getRowNum() - 1);
    }

    /**
     * 跳转到本行下方的第n行
     * @param num 向下跳过的行数
     * @return 指定行
     */
    public ExcelRow nextRow(int num) {
        return rowAt(this.row.getRowNum() + num);
    }

    /**
     * 下一行
     * @return 下一行
     */
    public ExcelRow nextRow() {
        return rowAt(this.row.getRowNum() + 1);
    }

    /**
     * 行高
     * @param num 行高的值
     * @return 本行
     */
    public ExcelRow height(int num) {
        this.row.setHeight((short) num);
        return this;
    }

    /**
     * 以指定Collection为基础，从本行开始循环生成行。
     * @param collection  Collection对象
     * @param func 行生成器
     * @return 从本行开始向下第Collection.size行。
     * @param <E> 数据对象的类型。
     */
    public <E> ExcelRow forOf(Collection<E> collection, RowIteratorFunction<E,ExcelCell> func) {
        ExcelRow cur = this;
        for (E e: collection) {
            func.accept(cur.cell(0),e);
            cur = cur.nextRow();
        }
        return cur;
    }

    /**
     * 跳转到指定列的Cell
     * @param col 列的index
     * @return 本行的第index列的Cell。
     */
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

    /**
     * 回到本行所在的Sheet。
     * @return 本Sheet。
     */
    public ExcelSheet backToSheet() {
        return sheet;
    }

}
