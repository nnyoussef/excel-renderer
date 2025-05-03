package lu.nyo.excel.renderer.utils;

import lu.nyo.excel.renderer.constantes.CssConstantes;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xssf.usermodel.XSSFColor;

import java.awt.*;

import static java.lang.Integer.parseInt;

public final class ColorUtils {
    private ColorUtils() throws InstantiationException {
        throw new InstantiationException();
    }

    public static XSSFColor getXSSFColorFromRgb(String rgb) {
        final XSSFColor xssfColor = new XSSFColor();
        xssfColor.setARGBHex(CssConstantes.HEX_WHITE);

        if (StringUtils.isEmpty(rgb)) {
            xssfColor.setARGBHex(CssConstantes.HEX_WHITE);
        } else if (rgb.contains("rgb")) {
            rgb = rgb.replace("rgb(", "").replace(")", "");
            String[] rgbArr = rgb.split(",");

            final int r = parseInt(rgbArr[0].trim().strip());
            final int g = parseInt(rgbArr[1].trim().strip());
            final int b = parseInt(rgbArr[2].trim().strip());

            final Color color = new Color(r, g, b);
            final String buf = Integer.toHexString(color.getRGB());
            final String hex = buf.substring(buf.length() - 6);
            xssfColor.setARGBHex(hex);
        }
        return xssfColor;
    }
}
