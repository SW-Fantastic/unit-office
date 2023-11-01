package org.swdc.offices.xlsx;

import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.swdc.offices.UnitUtils;

import java.awt.*;
import java.util.function.Consumer;

/**
 * 适用于XLSX（XSSF）格式的字体修改器。
 * @param <T> 创建字体修改器的对象类型
 */
public class ExcelFont<T> {

    private XSSFFont font;

    private Consumer<XSSFFont> callback;

    private T target;

    public ExcelFont(XSSFFont font,Consumer<XSSFFont> callback,T target) {
        this.font = font;
        this.callback = callback;
        this.target = target;
    }

    /**
     * 修改字体样式
     * @param name 字体名
     * @return 本对象
     */
    public ExcelFont<T> name(String name) {
        this.font.setFontName(name);
        return this;
    }

    /**
     * 修改字体大小
     * @param size 字体大小
     * @return 本对象
     */
    public ExcelFont<T> size(int size) {
        this.font.setFontHeightInPoints((short) size);
        return this;
    }

    /**
     * 字体加粗
     * @param val true = 加粗，false = 不加粗
     * @return 本对象
     */
    public ExcelFont<T> bold(boolean val) {
        this.font.setBold(val);
        return this;
    }

    /**
     * 修改删除线
     * @param val 有删除线 = true 无删除线 = false
     * @return 本对象
     */
    public ExcelFont<T> strikeout(boolean val) {
        this.font.setStrikeout(val);
        return this;
    }

    /**
     * 是否需要下划线
     * @param val true = 有下划线 ，false = 无下划线
     * @return 本对象
     */
    public ExcelFont<T> underline(boolean val) {
        this.font.setUnderline(val ? FontUnderline.SINGLE : FontUnderline.NONE);
        return this;
    }

    /**
     * 是否需要双下划线
     * @param val true = 右双下划线，false = 无双下划线
     * @return 本对象
     */
    public ExcelFont<T> doubleUnderline(boolean val) {
        this.font.setUnderline(val ? FontUnderline.DOUBLE : FontUnderline.NONE);
        return this;
    }

    /**
     * 修改字体颜色
     * @param color 颜色字符串
     * @return 本对象
     */
    public ExcelFont<T> color(String color) {
        if (color == null || color.isBlank()) {
            return this;
        }
        Color theColor = UnitUtils.fromString(color);
        if (theColor == null) {
            return this;
        }
        XSSFColor target = new XSSFColor(theColor,null);
        this.font.setColor(target);
        return this;
    }

    /**
     * 结束字体修改，返回上一层。
     * @return 创建字体修改器的对象。
     */
    public T back() {
        callback.accept(font);
        return this.target;
    }

}
