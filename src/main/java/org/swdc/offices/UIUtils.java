package org.swdc.offices;

import java.awt.*;

public class UIUtils {

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

}
