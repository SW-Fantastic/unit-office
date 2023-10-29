package org.swdc.offices.xlsx;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.swdc.offices.CellPresetFunction;
import org.swdc.offices.UIUtils;

import java.util.Date;

public class ExcelCell {

    private XSSFCell cell;

    private ExcelRow row;

    private XSSFCellStyle style;

    private XSSFFont font;

    public ExcelCell(ExcelRow row, XSSFCell cell) {
        this.row = row;
        this.cell = cell;
    }

    public ExcelCell type(CellType type) {
        this.cell.setCellType(type);
        return this;
    }

    public ExcelCell text(String text) {
        this.cell.setCellValue(text);
        return this;
    }

    public ExcelCell number(Double val) {
        this.cell.setCellValue(val);
        return this;
    }

    public ExcelCell date(Date date) {
        this.cell.setCellValue(date);
        return this;
    }

    public ExcelPicture<ExcelCell> picture() {
        return new ExcelPicture<>(cell.getSheet(),this)
                .position(
                        cell.getRowIndex(),
                        cell.getColumnIndex()
                );
    }

    public ExcelCell preset(CellPresetFunction preset) {
        return preset.accept(this);
    }

    private XSSFCellStyle getStyle() {
        if(this.style == null) {
            style = cell.getRow()
                    .getSheet()
                    .getWorkbook()
                    .createCellStyle();
        }
        return style;
    }

    private XSSFFont getFont() {
        if (this.font == null) {
            this.font = cell
                    .getRow()
                    .getSheet()
                    .getWorkbook()
                    .createFont();
        }
        return font;
    }

    public ExcelFont<ExcelCell> font() {
        return new ExcelFont<>(getFont(),f -> {
            CellStyle theStyle = getStyle();
            theStyle.setFont(f);
            cell.setCellStyle(style);
        },this);
    }

    public ExcelCell borderLeft(BorderStyle style) {
        XSSFCellStyle xsStyle = getStyle();
        xsStyle.setBorderLeft(style);
        return this;
    }

    public ExcelCell borderLeftColor(String color) {
        XSSFCellStyle xsStyle = getStyle();
        xsStyle.setLeftBorderColor(new XSSFColor(UIUtils.fromString(color),null));
        return this;
    }

    public ExcelCell borderRight(BorderStyle style) {
        XSSFCellStyle xsStyle = getStyle();
        xsStyle.setBorderRight(style);
        return this;
    }

    public ExcelCell borderRightColor(String color) {
        XSSFCellStyle xsStyle = getStyle();
        xsStyle.setRightBorderColor(new XSSFColor(UIUtils.fromString(color),null));
        return this;
    }

    public ExcelCell borderTop(BorderStyle style) {
        XSSFCellStyle xssStyle = getStyle();
        xssStyle.setBorderTop(style);
        return this;
    }

    public ExcelCell borderTopColor(String color) {
        XSSFCellStyle xsStyle = getStyle();
        xsStyle.setTopBorderColor(new XSSFColor(UIUtils.fromString(color),null));
        return this;
    }

    public ExcelCell borderBottom(BorderStyle style) {
        XSSFCellStyle xssStyle = getStyle();
        xssStyle.setBorderBottom(style);
        return this;
    }

    public ExcelCell borderBottomColor(String color) {
        XSSFCellStyle xsStyle = getStyle();
        xsStyle.setBottomBorderColor(new XSSFColor(UIUtils.fromString(color),null));
        return this;
    }

    public ExcelCell border(BorderStyle style) {

        XSSFCellStyle xssStyle = getStyle();

        xssStyle.setBorderBottom(style);
        xssStyle.setBorderRight(style);
        xssStyle.setBorderLeft(style);
        xssStyle.setBorderTop(style);

        return this;
    }

    public ExcelCell borderColor(String color) {

        XSSFColor theColor = new XSSFColor(UIUtils.fromString(color),null);
        XSSFCellStyle xssStyle = getStyle();
        xssStyle.setBottomBorderColor(theColor);
        xssStyle.setTopBorderColor(theColor);
        xssStyle.setLeftBorderColor(theColor);
        xssStyle.setRightBorderColor(theColor);

        cell.setCellStyle(xssStyle);
        return this;
    }

    // Aligns
    public ExcelCell align(HorizontalAlignment alignment) {

        XSSFCellStyle style = getStyle();

        style.setAlignment(alignment);
        cell.setCellStyle(style);
        return this;
    }

    public ExcelCell verticalAlignment(VerticalAlignment alignment){
        XSSFCellStyle style = getStyle();

        style.setVerticalAlignment(alignment);
        cell.setCellStyle(style);
        return this;
    }

    public ExcelCell alignVerticalCenter() {
        return verticalAlignment(VerticalAlignment.CENTER);
    }

    public ExcelCell alignVerticalTop() {
        return verticalAlignment(VerticalAlignment.TOP);
    }

    public ExcelCell alignVerticalBottom() {
        return verticalAlignment(VerticalAlignment.BOTTOM);
    }

    public ExcelCell alignCenter() {
        return align(HorizontalAlignment.CENTER);
    }

    public ExcelCell alignLeft() {
        return align(HorizontalAlignment.LEFT);
    }

    public ExcelCell alignRight() {
        return align(HorizontalAlignment.RIGHT);
    }

    public ExcelCell alignFill() {
        return align(HorizontalAlignment.FILL);
    }

    // Aligns - End

    // Positions

    public ExcelCell nextCell() {
        return this.row.cell(this.cell.getColumnIndex() + 1);
    }

    public ExcelCell prevCell() {
        if (this.cell.getColumnIndex() == 0) {
            throw new RuntimeException("this is already the first column");
        }
        return this.row.cell(this.cell.getColumnIndex() - 1);
    }

    public ExcelCell cellAt(int column) {
        if (column < 0) {
            throw new RuntimeException("invalid column");
        }
        return this.row.cell(column);
    }

    public ExcelRow backToRow() {
        return this.row;
    }

    // Positions End

}
