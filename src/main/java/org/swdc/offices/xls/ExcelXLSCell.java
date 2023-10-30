package org.swdc.offices.xls;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.swdc.offices.CellPresetFunction;
import org.swdc.offices.UIUtils;

import java.awt.Color;
import java.util.Date;
import java.util.function.Consumer;

/**
 * XLS Cell，
 * 用于XLS格式（HSSF）的单元格类。
 *
 * 本类用于为单元格填充数据，设定单元格的字体和格式等，
 * 同样，它可以为单元格添加图片。
 *
 */
public class ExcelXLSCell {

    private HSSFCell cell;

    private ExcelXLSRow row;

    private HSSFCellStyle style;

    private HSSFFont font;

    public ExcelXLSCell(ExcelXLSRow row, HSSFCell cell) {
        this.row = row;
        this.cell = cell;
    }

    /**
     * 设置单元格类型
     * @param type POI单元格类型
     * @return 本cell
     */
    public ExcelXLSCell type(CellType type) {
        this.cell.setCellType(type);
        return this;
    }

    /**
     * 在本单元格填写文本
     * @param text 单元格内容
     * @return 本cell
     */
    public ExcelXLSCell text(String text) {
        this.cell.setCellValue(text);
        return this;
    }

    /**
     * 在本单元格填写数值
     * @param val 单元格内容
     * @return 本cell
     */
    public ExcelXLSCell number(Double val) {
        this.cell.setCellValue(val);
        return this;
    }

    /**
     * 在本单元格填写日期
     * @param date 单元格内容
     * @return 本cell
     */
    public ExcelXLSCell date(Date date) {
        this.cell.setCellValue(date);
        return this;
    }

    /**
     * 在本单元格插入图片。
     * 请使用返回的对象进一步完成图片插入的操作。
     *
     * @return 本单元格的图片插入器。
     */
    public ExcelXLSPicture<ExcelXLSCell> picture() {
        return new ExcelXLSPicture<>(cell.getSheet(),this)
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
    public ExcelXLSCell preset(CellPresetFunction<ExcelXLSCell> preset) {
        if (preset == null){
            return this;
        }
        return preset.accept(this);
    }

    /**
     * 读取单元格样式，内部API，不给用户访问。
     * @return 单元格样式
     */
    private HSSFCellStyle getStyle() {
        if(this.style == null) {
            style = cell.getCellStyle();
        }
        return style;
    }

    /**
     * 读取或创建单元格字体，内部API，不给用户访问
     * @return 单元格字体
     */
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

    /**
     * 在工作簿的调色板读取或增加自定义的颜色，
     * 内部API，不给用户访问。
     *
     * @param color awt颜色
     * @return 单元格颜色
     */
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

    /**
     * 颜色处理方法，用于把颜色字符串转换为HSSFColor
     * 内部API，不给用户访问。
     *
     * @param color 颜色字符串
     * @param colorConsumer 颜色处理回调。
     * @return 本cell
     */
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

    /**
     * 修改字体样式
     * @return 本单元格字体修改器
     */
    public ExcelXLSFont<ExcelXLSCell> font() {
        return new ExcelXLSFont<>(cell.getSheet(),getFont(),(f) -> {
            CellStyle theStyle = getStyle();
            theStyle.setFont(f);
            cell.setCellStyle(style);
        },this);
    }

    /**
     * 修改左侧边框样式
     * @param style 样式
     * @return 本Cell
     */
    public ExcelXLSCell borderLeft(BorderStyle style) {
        HSSFCellStyle xsStyle = getStyle();
        xsStyle.setBorderLeft(style);
        cell.setCellStyle(xsStyle);
        return this;
    }

    /**
     * 修改左侧边框颜色
     * @param color color字符串
     * @return 本cell
     */
    public ExcelXLSCell borderLeftColor(String color) {
        return appendColor(color, c-> {
            HSSFCellStyle xsStyle = getStyle();
            xsStyle.setLeftBorderColor(c.getIndex());
            cell.setCellStyle(xsStyle);
        });
    }

    /**
     * 修改右侧边框样式
     * @param style 样式
     * @return 本Cell
     */
    public ExcelXLSCell borderRight(BorderStyle style) {
        HSSFCellStyle xsStyle = getStyle();
        xsStyle.setBorderRight(style);
        cell.setCellStyle(xsStyle);
        return this;
    }

    /**
     * 修改右侧边框颜色
     * @param color color字符串
     * @return 本cell
     */
    public ExcelXLSCell borderRightColor(String color) {
        HSSFCellStyle xsStyle = getStyle();
        return this.appendColor(color,c -> {
            xsStyle.setRightBorderColor(c.getIndex());
            cell.setCellStyle(xsStyle);
        });
    }


    /**
     * 修改上方边框样式
     * @param style 样式
     * @return 本Cell
     */
    public ExcelXLSCell borderTop(BorderStyle style) {
        HSSFCellStyle xssStyle = getStyle();
        xssStyle.setBorderTop(style);
        cell.setCellStyle(xssStyle);
        return this;
    }

    /**
     * 修改上方边框颜色
     * @param color color字符串
     * @return 本cell
     */
    public ExcelXLSCell borderTopColor(String color) {
        return appendColor(color, c-> {
            HSSFCellStyle xsStyle = getStyle();
            xsStyle.setTopBorderColor(c.getIndex());
            cell.setCellStyle(style);
        });
    }

    /**
     * 修改下方边框样式
     * @param style 样式
     * @return 本Cell
     */
    public ExcelXLSCell borderBottom(BorderStyle style) {
        HSSFCellStyle xssStyle = getStyle();
        xssStyle.setBorderBottom(style);
        cell.setCellStyle(xssStyle);
        return this;
    }

    /**
     * 修改下方边框颜色
     * @param color color字符串
     * @return 本cell
     */
    public ExcelXLSCell borderBottomColor(String color) {
       return appendColor(color, c -> {
           HSSFCellStyle xsStyle = getStyle();
           xsStyle.setBottomBorderColor(c.getIndex());
           cell.setCellStyle(xsStyle);
       });
    }

    /**
     * 修改所有边框样式
     * @param style 样式
     * @return 本Cell
     */
    public ExcelXLSCell border(BorderStyle style) {

        HSSFCellStyle xssStyle = getStyle();

        xssStyle.setBorderBottom(style);
        xssStyle.setBorderRight(style);
        xssStyle.setBorderLeft(style);
        xssStyle.setBorderTop(style);
        cell.setCellStyle(xssStyle);

        return this;
    }

    /**
     * 修改所有边框颜色
     * @param color color字符串
     * @return 本cell
     */
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

    /**
     * 修改水平对方式
     * @param alignment 水平对齐方式
     * @return 本cell
     */
    public ExcelXLSCell align(HorizontalAlignment alignment) {

        HSSFCellStyle style = getStyle();

        style.setAlignment(alignment);
        cell.setCellStyle(style);
        return this;
    }

    /**
     * 修改垂直对齐方式
     * @param alignment 垂直对齐方式
     * @return 本cell
     */
    public ExcelXLSCell verticalAlignment(VerticalAlignment alignment){
        HSSFCellStyle style = getStyle();

        style.setVerticalAlignment(alignment);
        cell.setCellStyle(style);
        return this;
    }

    /**
     * 垂直居中
     * @return 本cell
     */
    public ExcelXLSCell alignVerticalCenter() {
        return verticalAlignment(VerticalAlignment.CENTER);
    }

    /**
     * 垂直居顶
     * @return 本cell
     */
    public ExcelXLSCell alignVerticalTop() {
        return verticalAlignment(VerticalAlignment.TOP);
    }

    /**
     * 垂直居底
     * @return 本cell
     */
    public ExcelXLSCell alignVerticalBottom() {
        return verticalAlignment(VerticalAlignment.BOTTOM);
    }

    /**
     * 水平居中
     * @return 本cell
     */
    public ExcelXLSCell alignCenter() {
        return align(HorizontalAlignment.CENTER);
    }

    /**
     * 水平居左
     * @return 本cell
     */
    public ExcelXLSCell alignLeft() {
        return align(HorizontalAlignment.LEFT);
    }

    /**
     * 水平居右
     * @return 本cell
     */
    public ExcelXLSCell alignRight() {
        return align(HorizontalAlignment.RIGHT);
    }

    /**
     * 水平填充
     * @return 本cell
     */
    public ExcelXLSCell alignFill() {
        return align(HorizontalAlignment.FILL);
    }

    // Aligns - End

    // Positions

    /**
     * 下一个cell
     * @return 本Cell右侧的Cell，如果不存在会创建。
     */
    public ExcelXLSCell nextCell() {
        return this.row.cell(this.cell.getColumnIndex() + 1);
    }

    /**
     * 上一个Cell
     * @return 本Cell左侧的Cell，如果不存在会创建。
     */
    public ExcelXLSCell prevCell() {
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
    public ExcelXLSCell cellAt(int column) {
        if (column < 0) {
            throw new RuntimeException("invalid column");
        }
        return this.row.cell(column);
    }

    /**
     * 返回Cell所在的Row
     * @return 本行。
     */
    public ExcelXLSRow backToRow() {
        return this.row;
    }

    // Positions End

}
