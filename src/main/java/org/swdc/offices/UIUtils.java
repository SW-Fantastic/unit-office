package org.swdc.offices;

import java.awt.*;

public class UIUtils {

    public static Color fromString(String colorStr) {
        colorStr = colorStr.toLowerCase();
        if (colorStr.startsWith("#")) {
            // hex string
            colorStr = colorStr.substring(1);
            if (colorStr.length() == 3) {
                // RGB
                return new Color(
                        Integer.parseInt(colorStr.substring(0,1).repeat(2),16),
                        Integer.parseInt(colorStr.substring(1,2).repeat(2),16),
                        Integer.parseInt(colorStr.substring(2).repeat(2),16)
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
