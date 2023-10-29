package org.swdc.offices.xlsx;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.swdc.offices.CellPresetFunction;
import org.swdc.offices.UIUtils;

import java.awt.Color;
import java.util.Date;
import java.util.function.Consumer;

/**
 * XLSX Cell，
 * 用于XLSX格式（XSSF）的单元格类。
 *
 * 本类用于为单元格填充数据，设定单元格的字体和格式等，
 * 同样，它可以为单元格添加图片。
 *
 */
public class ExcelCell {

    private XSSFCell cell;

    private ExcelRow row;

    private XSSFCellStyle style;

    private XSSFFont font;

    public ExcelCell(ExcelRow row, XSSFCell cell) {
        this.row = row;
        this.cell = cell;
    }

    /**
     * 设置单元格类型
     * @param type POI单元格类型
     * @return 本cell
     */
    public ExcelCell type(CellType type) {
        this.cell.setCellType(type);
        return this;
    }

    /**
     * 在本单元格填写文本
     * @param text 单元格内容
     * @return 本cell
     */
    public ExcelCell text(String text) {
        this.cell.setCellValue(text);
        return this;
    }

    /**
     * 在本单元格填写数值
     * @param val 单元格内容
     * @return 本cell
     */
    public ExcelCell number(Double val) {
        this.cell.setCellValue(val);
        return this;
    }

    /**
     * 在本单元格填写日期
     * @param date 单元格内容
     * @return 本cell
     */
    public ExcelCell date(Date date) {
        this.cell.setCellValue(date);
        return this;
    }

    /**
     * 在本单元格插入图片。
     * 请使用返回的对象进一步完成图片插入的操作。
     *
     * @return 本单元格的图片插入器。
     */
    public ExcelPicture<ExcelCell> picture() {
        return new ExcelPicture<>(cell.getSheet(),this)
                .position(
                        cell.getRowIndex(),
                        cell.getColumnIndex()
                );
    }

    /**
     * 在本单元格使用预设样式。
     *
     * @param preset 预设样式函数。
     * @return 本cell
     */
    public ExcelCell preset(CellPresetFunction<ExcelCell> preset) {
        return preset.accept(this);
    }

    /**
     * 读取单元格样式，内部API，不给用户访问。
     * @return 单元格样式
     */
    private XSSFCellStyle getStyle() {
        if(this.style == null) {
            style = cell.getRow()
                    .getSheet()
                    .getWorkbook()
                    .createCellStyle();
        }
        return style;
    }

    /**
     * 读取或创建单元格字体，内部API，不给用户访问
     * @return 单元格字体
     */
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

    /**
     * 修改字体样式
     * @return 本单元格字体修改器
     */
    public ExcelFont<ExcelCell> font() {
        return new ExcelFont<>(getFont(),f -> {
            CellStyle theStyle = getStyle();
            theStyle.setFont(f);
            cell.setCellStyle(style);
        },this);
    }

    /**
     * 颜色处理方法，用于把颜色字符串转换为HSSFColor
     * 内部API，不给用户访问。
     *
     * @param color 颜色字符串
     * @param colorConsumer 颜色处理回调。
     * @return 本cell
     */
    private ExcelCell appendColor(String color, Consumer<Color> colorConsumer) {
        if (color == null || color.isBlank()) {
            return this;
        }
        Color awtColor = UIUtils.fromString(color);
        if (awtColor == null) {
            return this;
        }
        colorConsumer.accept(awtColor);
        return this;
    }


    /**
     * 修改左侧边框样式
     * @param style 样式
     * @return 本Cell
     */
    public ExcelCell borderLeft(BorderStyle style) {
        XSSFCellStyle xsStyle = getStyle();
        xsStyle.setBorderLeft(style);
        return this;
    }

    /**
     * 修改左侧边框颜色
     * @param color color字符串
     * @return 本cell
     */
    public ExcelCell borderLeftColor(String color) {
        return appendColor(color,c -> {
            XSSFCellStyle xsStyle = getStyle();
            xsStyle.setLeftBorderColor(new XSSFColor(c,null));
        });
    }

    /**
     * 修改右侧边框样式
     * @param style 样式
     * @return 本Cell
     */
    public ExcelCell borderRight(BorderStyle style) {
        XSSFCellStyle xsStyle = getStyle();
        xsStyle.setBorderRight(style);
        return this;
    }

    /**
     * 修改右侧边框颜色
     * @param color color字符串
     * @return 本cell
     */
    public ExcelCell borderRightColor(String color) {
        return appendColor(color, c -> {
            XSSFCellStyle xsStyle = getStyle();
            xsStyle.setRightBorderColor(new XSSFColor(c,null));
        });
    }

    /**
     * 修改上方边框样式
     * @param style 样式
     * @return 本Cell
     */
    public ExcelCell borderTop(BorderStyle style) {
        XSSFCellStyle xssStyle = getStyle();
        xssStyle.setBorderTop(style);
        return this;
    }

    /**
     * 修改上方边框颜色
     * @param color color字符串
     * @return 本cell
     */
    public ExcelCell borderTopColor(String color) {
        return appendColor(color, c -> {
            XSSFCellStyle xsStyle = getStyle();
            xsStyle.setTopBorderColor(new XSSFColor(c,null));
        });
    }

    /**
     * 修改下方边框样式
     * @param style 样式
     * @return 本Cell
     */
    public ExcelCell borderBottom(BorderStyle style) {
        XSSFCellStyle xssStyle = getStyle();
        xssStyle.setBorderBottom(style);
        return this;
    }

    /**
     * 修改下方边框颜色
     * @param color color字符串
     * @return 本cell
     */
    public ExcelCell borderBottomColor(String color) {
        XSSFCellStyle xsStyle = getStyle();
        xsStyle.setBottomBorderColor(new XSSFColor(UIUtils.fromString(color),null));
        return this;
    }

    /**
     * 修改所有边框样式
     * @param style 样式
     * @return 本Cell
     */
    public ExcelCell border(BorderStyle style) {

        XSSFCellStyle xssStyle = getStyle();

        xssStyle.setBorderBottom(style);
        xssStyle.setBorderRight(style);
        xssStyle.setBorderLeft(style);
        xssStyle.setBorderTop(style);

        return this;
    }

    /**
     * 修改所有边框颜色
     * @param color color字符串
     * @return 本cell
     */
    public ExcelCell borderColor(String color) {

        return appendColor(color , c -> {
            XSSFColor theColor = new XSSFColor(c,null);
            XSSFCellStyle xssStyle = getStyle();
            xssStyle.setBottomBorderColor(theColor);
            xssStyle.setTopBorderColor(theColor);
            xssStyle.setLeftBorderColor(theColor);
            xssStyle.setRightBorderColor(theColor);
            cell.setCellStyle(xssStyle);
        });
    }

    // Aligns
    /**
     * 修改水平对方式
     * @param alignment 水平对齐方式
     * @return 本cell
     */
    public ExcelCell align(HorizontalAlignment alignment) {

        XSSFCellStyle style = getStyle();

        style.setAlignment(alignment);
        cell.setCellStyle(style);
        return this;
    }

    /**
     * 修改垂直对齐方式
     * @param alignment 垂直对齐方式
     * @return 本cell
     */
    public ExcelCell verticalAlignment(VerticalAlignment alignment){
        XSSFCellStyle style = getStyle();

        style.setVerticalAlignment(alignment);
        cell.setCellStyle(style);
        return this;
    }

    /**
     * 垂直居中
     * @return 本cell
     */
    public ExcelCell alignVerticalCenter() {
        return verticalAlignment(VerticalAlignment.CENTER);
    }

    /**
     * 垂直居顶
     * @return 本cell
     */
    public ExcelCell alignVerticalTop() {
        return verticalAlignment(VerticalAlignment.TOP);
    }

    /**
     * 垂直居底
     * @return 本cell
     */
    public ExcelCell alignVerticalBottom() {
        return verticalAlignment(VerticalAlignment.BOTTOM);
    }

    /**
     * 水平居中
     * @return 本cell
     */
    public ExcelCell alignCenter() {
        return align(HorizontalAlignment.CENTER);
    }

    /**
     * 水平居左
     * @return 本cell
     */
    public ExcelCell alignLeft() {
        return align(HorizontalAlignment.LEFT);
    }

    /**
     * 水平居右
     * @return 本cell
     */
    public ExcelCell alignRight() {
        return align(HorizontalAlignment.RIGHT);
    }

    /**
     * 水平填充
     * @return 本cell
     */
    public ExcelCell alignFill() {
        return align(HorizontalAlignment.FILL);
    }

    // Aligns - End

    // Positions

    /**
     * 下一个cell
     * @return 本Cell右侧的Cell，如果不存在会创建。
     */
    public ExcelCell nextCell() {
        return this.row.cell(this.cell.getColumnIndex() + 1);
    }

    /**
     * 上一个Cell
     * @return 本Cell左侧的Cell，如果不存在会创建。
     */
    public ExcelCell prevCell() {
        if (this.cell.getColumnIndex() == 0) {
            throw new RuntimeException("this is already the first column");
        }
        return this.row.cell(this.cell.getColumnIndex() - 1);
    }

    /**
     * 指定列的Cell
     * @param column 列index
     * @return 此列的Cell
     */
    public ExcelCell cellAt(int column) {
        if (column < 0) {
            throw new RuntimeException("invalid column");
        }
        return this.row.cell(column);
    }

    /**
     * 返回Cell所在的Row
     * @return 本行。
     */
    public ExcelRow backToRow() {
        return this.row;
    }

    // Positions End

}
