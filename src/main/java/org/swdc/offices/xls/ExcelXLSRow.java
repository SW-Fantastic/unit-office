package org.swdc.offices.xls;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.swdc.offices.CellPresetFunction;
import org.swdc.offices.RowIteratorFunction;
import org.swdc.offices.xlsx.ExcelSheet;

import java.util.Collection;

/**
 * XLS（HSSF）表中的一行。
 */
public class ExcelXLSRow {

    private HSSFRow row;

    private ExcelXLSSheet sheet;

    private CellPresetFunction preset;

    public ExcelXLSRow(ExcelXLSSheet sheet, HSSFRow row) {
        this.sheet = sheet;
        this.row = row;
    }

    /**
     * 为本行全部的Cell设置预设样式
     * @param preset 预设样式函数
     * @return 本行
     */
    public ExcelXLSRow presetCell(CellPresetFunction preset) {
        this.preset = preset;
        return this;
    }

    /**
     * 到指定的行，如果不存在则会创建
     * @param idx 指定行的index。
     * @return 指定行
     */
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

    /**
     * 跳转到上方的第n行
     * @param num 向上跳过的行数
     * @return 指定行
     */
    public ExcelXLSRow prevRow(int num) {
        return rowAt(this.row.getRowNum() - num);
    }

    /**
     * 上一行
     * @return 本行的上一行
     */
    public ExcelXLSRow prevRow() {
        return rowAt(this.row.getRowNum() - 1);
    }

    /**
     * 跳转到本行下方的第n行
     * @param num 向下跳过的行数
     * @return 指定行
     */
    public ExcelXLSRow nextRow(int num) {
        return rowAt(this.row.getRowNum() + num);
    }

    /**
     * 下一行
     * @return 下一行
     */
    public ExcelXLSRow nextRow() {
        return rowAt(this.row.getRowNum() + 1);
    }

    /**
     * 行高
     * @param num 行高的值
     * @return 本行
     */
    public ExcelXLSRow height(int num) {
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
    public <E> ExcelXLSRow forOf(Collection<E> collection, RowIteratorFunction<E,ExcelXLSCell> func) {
        ExcelXLSRow cur = this;
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

    /**
     * 回到本行所在的Sheet。
     * @return 本Sheet。
     */
    public ExcelXLSSheet backToSheet() {
        return sheet;
    }


}
