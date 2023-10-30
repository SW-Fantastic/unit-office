package org.swdc.offices.xls;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Font;
import org.swdc.offices.UIUtils;

import java.awt.Color;
import java.util.function.Consumer;

/**
 * 适用于XLS（HSSF）格式的字体修改器。
 * @param <T> 创建字体修改器的对象类型
 */
public class ExcelXLSFont<T> {

    private HSSFFont font;

    private Consumer<HSSFFont> callback;

    private T target;

    private HSSFSheet sheet;

    public ExcelXLSFont(HSSFSheet sheet, HSSFFont font, Consumer<HSSFFont> callback, T target) {
        this.font = font;
        this.callback = callback;
        this.target = target;
        this.sheet = sheet;
    }

    /**
     * 修改字体样式
     * @param name 字体名
     * @return 本对象
     */
    public ExcelXLSFont<T> name(String name) {
        this.font.setFontName(name);
        return this;
    }

    /**
     * 修改删除线
     * @param val 有删除线 = true 无删除线 = false
     * @return 本对象
     */
    public ExcelXLSFont<T> strikeout(boolean val) {
        font.setStrikeout(val);
        return this;
    }

    /**
     * 修改字体大小
     * @param size 字体大小
     * @return 本对象
     */
    public ExcelXLSFont<T> size(int size) {
        this.font.setFontHeightInPoints((short) size);
        return this;
    }

    /**
     * 字体加粗
     * @param val true = 加粗，false = 不加粗
     * @return 本对象
     */
    public ExcelXLSFont<T> bold(boolean val) {
        this.font.setBold(val);
        return this;
    }

    /**
     * 是否需要下划线
     * @param val true = 有下划线 ，false = 无下划线
     * @return 本对象
     */
    public ExcelXLSFont<T> underline(boolean val) {
        this.font.setUnderline(Font.U_SINGLE);
        return this;
    }

    /**
     * 是否需要双下划线
     * @param val true = 右双下划线，false = 无双下划线
     * @return 本对象
     */
    public ExcelXLSFont<T> doubleUnderline(boolean val) {
        this.font.setUnderline(Font.U_DOUBLE);
        return this;
    }


    /**
     * 修改字体颜色
     * @param color 颜色字符串
     * @return 本对象
     */
    public ExcelXLSFont<T> color(String color) {
        if (color == null || color.isBlank()) {
            return this;
        }
        Color theColor = UIUtils.fromString(color);
        if (theColor == null) {
            return this;
        }
        HSSFColor realColor = UIUtils.prepareHSSFColor(
                sheet.getWorkbook(),
                theColor
        );
        if (realColor == null) {
            return this;
        }
        this.font.setColor(realColor.getIndex());
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
