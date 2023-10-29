package org.swdc.offices.xls;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Font;
import org.swdc.offices.UIUtils;

import java.awt.Color;
import java.util.function.Consumer;

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

    public ExcelXLSFont<T> name(String name) {
        this.font.setFontName(name);
        return this;
    }

    public ExcelXLSFont<T> size(int size) {
        this.font.setFontHeightInPoints((short) size);
        return this;
    }

    public ExcelXLSFont<T> bold(boolean val) {
        this.font.setBold(val);
        return this;
    }

    public ExcelXLSFont<T> underline(boolean val) {
        this.font.setUnderline(Font.U_SINGLE);
        return this;
    }

    public ExcelXLSFont<T> doubleUnderline(boolean val) {
        this.font.setUnderline(Font.U_DOUBLE);
        return this;
    }

    private HSSFColor addColor(Color color) {
        HSSFPalette palette = sheet.getWorkbook().getCustomPalette();
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

    public ExcelXLSFont<T> color(String color) {
        if (color == null || color.isBlank()) {
            return this;
        }
        Color theColor = UIUtils.fromString(color);
        if (theColor == null) {
            return this;
        }
        HSSFColor realColor = addColor(theColor);
        if (realColor == null) {
            return this;
        }
        this.font.setColor(realColor.getIndex());
        return this;
    }

    public T back() {
        callback.accept(font);
        return this.target;
    }

}
