package org.swdc.offices.xlsx;

import org.apache.poi.ss.usermodel.ShapeTypes;
import org.apache.poi.xssf.usermodel.*;
import org.swdc.offices.UnitUtils;

import java.awt.*;

public class ExcelShape<T> {

    private T target;

    private XSSFSheet sheet;

    private XSSFClientAnchor anchor;

    private XSSFSimpleShape shape;

    public ExcelShape(XSSFSheet sheet, T target) {
        this.target = target;
        this.sheet = sheet;
        XSSFWorkbook workbook = sheet.getWorkbook();
        XSSFCreationHelper helper = workbook.getCreationHelper();
        anchor = helper.createClientAnchor();

        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        shape = drawing.createSimpleShape(anchor);
    }

    /**
     * 指定形状的类型
     * @param type POI的形状类型
     * @return 本对象
     * @see org.apache.poi.ss.usermodel.ShapeTypes
     */
    public ExcelShape<T> type(int type) {
        shape.setShapeType(type);
        return this;
    }

    /**
     * 创建为矩形
     * @return 本对象
     */
    public ExcelShape<T> rect() {
        return type(ShapeTypes.RECT);
    }

    /**
     * 创建为圆角矩形
     * @return 本对象
     */
    public ExcelShape<T> roundedRect() {
        return type(ShapeTypes.ROUND_RECT);
    }

    /**
     * 创建为椭圆形
     * @return 本对象
     */
    public ExcelShape<T> ellipse() {
        return type(ShapeTypes.ELLIPSE);
    }


    /**
     * 指定本形状的前景色
     * @param color 颜色字符串
     * @return 本对象
     */
    public ExcelShape<T> color(String color) {
        if (color == null || color.isEmpty()) {
            return this;
        }
        Color awtColor = UnitUtils.fromString(color);
        if (awtColor == null) {
            return this;
        }
        shape.setLineStyleColor(awtColor.getRed(),awtColor.getGreen(),awtColor.getBlue());
        return this;
    }

    /**
     * 指定本形状的背景色
     * @param color 颜色字符串
     * @return 本对象
     */
    public ExcelShape<T> background(String color) {
        if (color == null || color.isEmpty()) {
            shape.setNoFill(true);
            return this;
        }
        Color awtColor = UnitUtils.fromString(color);
        if (awtColor == null) {
            shape.setNoFill(true);
            return this;
        }
        shape.setNoFill(false);
        shape.setFillColor(awtColor.getRed(),awtColor.getGreen(),awtColor.getBlue());
        return this;
    }

    /**
     * 指定形状的位置
     * @param rowBegin 形状将会从此行开始
     * @param colBegin 形状将会从此列开始
     * @param rowEnd 形状将会在此行结束
     * @param colEnd 形状将会在此列结束
     * @return 本对象
     */
    public ExcelShape<T> position(int rowBegin, int colBegin, int rowEnd, int colEnd) {
        anchor.setCol1(colBegin);
        anchor.setCol2(colEnd);
        anchor.setRow1(rowBegin);
        anchor.setRow2(rowEnd);
        return this;
    }

    /**
     * 修改形状插入的单元格位置
     * @param row 形状将会出现在此行
     * @param col 形状将会出现在此列
     * @return 本对象。
     */
    public ExcelShape<T> position(int row, int col) {
        return position(row,col,row,col);
    }

    /**
     * 跨行跨列处理
     * @param rowSpan 形状的纵向跨行数
     * @param colSpan 形状的横向跨列数
     * @return 本对象
     */
    public ExcelShape<T> cross(int rowSpan, int colSpan) {
        return position(
                anchor.getRow1(),
                anchor.getCol1(),
                anchor.getRow1() + rowSpan,
                anchor.getCol1() + colSpan
        );
    }

    public T back() {
        return target;
    }

}
