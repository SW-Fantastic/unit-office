package org.swdc.offices.xls;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.swdc.offices.CellPresetFunction;
import org.swdc.offices.UIUtils;

import java.awt.Color;
import java.util.Date;
import java.util.function.Consumer;

public class ExcelXLSCell {

    private HSSFCell cell;

    private ExcelXLSRow row;

    private HSSFCellStyle style;

    private HSSFFont font;

    public ExcelXLSCell(ExcelXLSRow row, HSSFCell cell) {
        this.row = row;
        this.cell = cell;
    }

    public ExcelXLSCell type(CellType type) {
        this.cell.setCellType(type);
        return this;
    }

    public ExcelXLSCell text(String text) {
        this.cell.setCellValue(text);
        return this;
    }

    public ExcelXLSCell number(Double val) {
        this.cell.setCellValue(val);
        return this;
    }

    public ExcelXLSCell date(Date date) {
        this.cell.setCellValue(date);
        return this;
    }

    public ExcelXLSPicture<ExcelXLSCell> picture() {
        return new ExcelXLSPicture<>(cell.getSheet(),this)
                .position(
                        cell.getRowIndex(),
                        cell.getColumnIndex()
                );
    }

    public ExcelXLSCell preset(CellPresetFunction<ExcelXLSCell> preset) {
        return preset.accept(this);
    }

    private HSSFCellStyle getStyle() {
        if(this.style == null) {
            style = cell.getRow()
                    .getSheet()
                    .getWorkbook()
                    .createCellStyle();
        }
        return style;
    }

    private HSSFFont getFont() {
        if (this.font == null) {
            this.font = cell
                    .getRow()
                    .getSheet()
                    .getWorkbook()
                    .createFont();
        }
        return font;
    }

    private HSSFColor addColor(Color color) {
        HSSFPalette palette = cell.getSheet().getWorkbook().getCustomPalette();
        HSSFColor target = palette.findColor(
                (byte) color.getRed(),
                (byte) color.getGreen(),
                (byte) color.getBlue()
        );
        if (target != null) {
            return target;
        }
        return palette.addColor(
                (byte) color.getRed(),
                (byte) color.getGreen(),
                (byte) color.getBlue()
        );
    }

    private ExcelXLSCell appendColor(String color, Consumer<HSSFColor> colorConsumer) {
        if (color == null || color.isBlank()) {
            return this;
        }
        Color awtColor = UIUtils.fromString(color);
        if (awtColor == null) {
            return this;
        }
        HSSFColor realColor = addColor(awtColor);
        if (realColor == null) {
            return this;
        }
        colorConsumer.accept(realColor);

        return this;
    }

    public ExcelXLSFont<ExcelXLSCell> font() {
        return new ExcelXLSFont<>(cell.getSheet(),getFont(),(f) -> {
            CellStyle theStyle = getStyle();
            theStyle.setFont(f);
            cell.setCellStyle(style);
        },this);
    }

    public ExcelXLSCell borderLeft(BorderStyle style) {
        HSSFCellStyle xsStyle = getStyle();
        xsStyle.setBorderLeft(style);
        return this;
    }

    public ExcelXLSCell borderLeftColor(String color) {
        return appendColor(color, c-> {
            HSSFCellStyle xsStyle = getStyle();
            xsStyle.setLeftBorderColor(c.getIndex());
        });
    }

    public ExcelXLSCell borderRight(BorderStyle style) {
        HSSFCellStyle xsStyle = getStyle();
        xsStyle.setBorderRight(style);
        return this;
    }

    public ExcelXLSCell borderRightColor(String color) {
        HSSFCellStyle xsStyle = getStyle();
        return this.appendColor(color,c -> {
            xsStyle.setRightBorderColor(c.getIndex());
        });
    }

    public ExcelXLSCell borderTop(BorderStyle style) {
        HSSFCellStyle xssStyle = getStyle();
        xssStyle.setBorderTop(style);
        return this;
    }

    public ExcelXLSCell borderTopColor(String color) {
        return appendColor(color, c-> {
            HSSFCellStyle xsStyle = getStyle();
            xsStyle.setTopBorderColor(c.getIndex());
        });
    }

    public ExcelXLSCell borderBottom(BorderStyle style) {
        HSSFCellStyle xssStyle = getStyle();
        xssStyle.setBorderBottom(style);
        return this;
    }

    public ExcelXLSCell borderBottomColor(String color) {
       return appendColor(color, c -> {
           HSSFCellStyle xsStyle = getStyle();
           xsStyle.setBottomBorderColor(c.getIndex());
       });
    }

    public ExcelXLSCell border(BorderStyle style) {

        HSSFCellStyle xssStyle = getStyle();

        xssStyle.setBorderBottom(style);
        xssStyle.setBorderRight(style);
        xssStyle.setBorderLeft(style);
        xssStyle.setBorderTop(style);

        return this;
    }

    public ExcelXLSCell borderColor(String color) {

       return appendColor(color, c-> {
            HSSFCellStyle xssStyle = getStyle();
            xssStyle.setBottomBorderColor(c.getIndex());
            xssStyle.setTopBorderColor(c.getIndex());
            xssStyle.setLeftBorderColor(c.getIndex());
            xssStyle.setRightBorderColor(c.getIndex());
            cell.setCellStyle(xssStyle);
        });

    }

    // Aligns
    public ExcelXLSCell align(HorizontalAlignment alignment) {

        HSSFCellStyle style = getStyle();

        style.setAlignment(alignment);
        cell.setCellStyle(style);
        return this;
    }

    public ExcelXLSCell verticalAlignment(VerticalAlignment alignment){
        HSSFCellStyle style = getStyle();

        style.setVerticalAlignment(alignment);
        cell.setCellStyle(style);
        return this;
    }

    public ExcelXLSCell alignVerticalCenter() {
        return verticalAlignment(VerticalAlignment.CENTER);
    }

    public ExcelXLSCell alignVerticalTop() {
        return verticalAlignment(VerticalAlignment.TOP);
    }

    public ExcelXLSCell alignVerticalBottom() {
        return verticalAlignment(VerticalAlignment.BOTTOM);
    }

    public ExcelXLSCell alignCenter() {
        return align(HorizontalAlignment.CENTER);
    }

    public ExcelXLSCell alignLeft() {
        return align(HorizontalAlignment.LEFT);
    }

    public ExcelXLSCell alignRight() {
        return align(HorizontalAlignment.RIGHT);
    }

    public ExcelXLSCell alignFill() {
        return align(HorizontalAlignment.FILL);
    }

    // Aligns - End

    // Positions

    public ExcelXLSCell nextCell() {
        return this.row.cell(this.cell.getColumnIndex() + 1);
    }

    public ExcelXLSCell prevCell() {
        if (this.cell.getColumnIndex() == 0) {
            throw new RuntimeException("this is already the first column");
        }
        return this.row.cell(this.cell.getColumnIndex() - 1);
    }

    public ExcelXLSCell cellAt(int column) {
        if (column < 0) {
            throw new RuntimeException("invalid column");
        }
        return this.row.cell(column);
    }

    public ExcelXLSRow backToRow() {
        return this.row;
    }

    // Positions End

}
