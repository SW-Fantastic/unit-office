package org.swdc.offices;

import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.swdc.offices.generator.GenerateFunction;
import org.swdc.offices.generator.GeneratorStage;
import org.swdc.offices.generator.PipedGenerationContext;

import java.awt.*;
import java.lang.reflect.Method;
import java.lang.reflect.Parameter;
import java.util.ArrayList;
import java.util.List;

public class UnitUtils {


    public static <S> List<GenerateFunction<S>> extractGenerateStages(Object strategy, Class<S> sheetType) {

        Class currType = strategy.getClass();
        List<GenerateFunction<S>> functions = new ArrayList<>();

        while (currType != Object.class) {
            Method[] methods = currType.getDeclaredMethods();
            for (Method m: methods) {
                GeneratorStage stage = m.getAnnotation(GeneratorStage.class);
                if (stage == null) {
                    continue;
                }
                Parameter[] params = m.getParameters();
                if (params.length == 2 && params[0].getType().equals(PipedGenerationContext.class) && params[1].getType().equals(sheetType)) {
                    m.setAccessible(true);
                    GenerateFunction<S> function = (ctx, sheet) -> {
                        try {
                            m.invoke(strategy,ctx,sheet);
                        } catch (Throwable t) {
                            throw new RuntimeException(t);
                        }
                    };
                    functions.add(function);
                }
            }
            currType = currType.getSuperclass();
        }

        return functions;
    }


    /**
     * 工具方法，用于将字符串进行重复（用于兼容Java8）
     * @param str string
     * @param times 重复的次数
     * @return 重复的结果
     */
    private static String repeatStr(String str, int times) {
        StringBuilder sb = new StringBuilder();
        for (int idx = 0; idx < times; idx ++) {
            sb.append(str);
        }
        return sb.toString();
    }

    /**
     * 常见的String转AWT Color的方法。
     * @param colorStr Color字符串，
     *                 支持三位Hex，例如：#FFF，
     *                 支持六位Hex，例如#CECECE，
     *                 支持rgb和rgba表达式，例如rgb(0,0,0)
     * @return awt Color
     */
    public static Color fromString(String colorStr) {
        colorStr = colorStr.toLowerCase();
        if (colorStr.startsWith("#")) {
            // hex string
            colorStr = colorStr.substring(1);
            if (colorStr.length() == 3) {
                // RGB
                return new Color(
                        Integer.parseInt(repeatStr(colorStr.substring(0,1),2),16),
                        Integer.parseInt(repeatStr(colorStr.substring(1,2),2),16),
                        Integer.parseInt(repeatStr(colorStr.substring(2),2),16)
                );
            } else if (colorStr.length() == 6) {
                // 两位RGB
                return new Color(
                        Integer.parseInt(colorStr.substring(0, 2), 16),
                        Integer.parseInt(colorStr.substring(2, 4), 16),
                        Integer.parseInt(colorStr.substring(4, 6), 16)
                );
            } else if (colorStr.length() == 8) {
                // 两位RGBA
                return new Color(
                        Integer.parseInt(colorStr.substring(0, 2), 16),
                        Integer.parseInt(colorStr.substring(2, 4), 16),
                        Integer.parseInt(colorStr.substring(4, 6), 16),
                        Integer.parseInt(colorStr.substring(6, 8), 16)
                );
            }
        } else if (colorStr.startsWith("rgb")) {
            if (colorStr.startsWith("rgb(")) {
                colorStr = colorStr.replace("rgb(","")
                        .replace(")","");
                String[] rgb = colorStr.split(",");
                return new Color(
                        Integer.parseInt(rgb[0]),
                        Integer.parseInt(rgb[1]),
                        Integer.parseInt(rgb[2])
                );
            } else if (colorStr.startsWith("rgba(")) {
                colorStr = colorStr.replace("rgba(","")
                        .replace(")","");
                String[] rgba = colorStr.split(",");
                String a = rgba[3];
                if (a.indexOf('.') > 0) {
                    double alpha = Double.parseDouble(a);
                    int intAlpha = (int)(alpha * 255);
                    a = "" + intAlpha;
                }
                return new Color(
                        Integer.parseInt(rgba[0]),
                        Integer.parseInt(rgba[1]),
                        Integer.parseInt(rgba[2]),
                        Integer.parseInt(a)
                );
            }
        }
        return null;
    }

    /**
     * 在工作簿的调色板读取或增加自定义的颜色，
     * 内部API，不给用户访问。
     *
     * @param color awt颜色
     * @return 单元格颜色
     */
    public static HSSFColor prepareHSSFColor(HSSFWorkbook workbook, Color color) {
        HSSFPalette palette = workbook.getCustomPalette();
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


}
