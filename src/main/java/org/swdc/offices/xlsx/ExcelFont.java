package org.swdc.offices.xlsx;

import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.swdc.offices.UIUtils;

import java.awt.*;
import java.util.function.Consumer;

public class ExcelFont<T> {

    private XSSFFont font;

    private Consumer<XSSFFont> callback;

    private T target;

    public ExcelFont(XSSFFont font,Consumer<XSSFFont> callback,T target) {
        this.font = font;
        this.callback = callback;
        this.target = target;
    }

    public ExcelFont<T> name(String name) {
        this.font.setFontName(name);
        return this;
    }

    public ExcelFont<T> size(int size) {
        this.font.setFontHeightInPoints((short) size);
        return this;
    }

    public ExcelFont<T> bold(boolean val) {
        this.font.setBold(val);
        return this;
    }

    public ExcelFont<T> underline(boolean val) {
        this.font.setUnderline(val ? FontUnderline.SINGLE : FontUnderline.NONE);
        return this;
    }

    public ExcelFont<T> doubleUnderline(boolean val) {
        this.font.setUnderline(val ? FontUnderline.DOUBLE : FontUnderline.NONE);
        return this;
    }

    public ExcelFont<T> color(String color) {
        Color theColor = UIUtils.fromString(color);
        XSSFColor target = new XSSFColor(theColor,null);
        this.font.setColor(target);
        return this;
    }

    public T back() {
        callback.accept(font);
        return this.target;
    }

}
