package lu.nyo.excel.renderer.styleprocessor;

import lu.nyo.excel.renderer.constantes.CssConstantes;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

import java.util.Arrays;
import java.util.Map;
import java.util.function.BiConsumer;

import static java.util.stream.Collectors.joining;
import static org.apache.commons.lang3.StringUtils.EMPTY;
import static org.apache.commons.lang3.StringUtils.SPACE;
import static org.apache.poi.ss.usermodel.BorderStyle.NONE;

public class BorderProcessor implements BiConsumer<Map<String, String>, XSSFCellStyle> {
    @Override
    public void accept(Map<String, String> cssProperties, XSSFCellStyle xssfCellStyle) {
        cssProperties.forEach((prop, value) -> {
            if (prop.equals(CssConstantes.BORDER)) {
                final XSSFColor xssfColor = getBorderColor(value);
                xssfCellStyle.setBorderBottom(getBorderStyle(value));
                xssfCellStyle.setBottomBorderColor(xssfColor);

                xssfCellStyle.setBorderTop(getBorderStyle(value));
                xssfCellStyle.setTopBorderColor(xssfColor);

                xssfCellStyle.setBorderLeft(getBorderStyle(value));
                xssfCellStyle.setLeftBorderColor(xssfColor);

                xssfCellStyle.setBorderRight(getBorderStyle(value));
                xssfCellStyle.setRightBorderColor(xssfColor);

            } else if (prop.contains(CssConstantes.BORDER_DASH)) {
                final String side = prop.split(CssConstantes.DASH)[1];
                final XSSFColor xssfColor = getBorderColor(value);
                switch (side) {
                    case CssConstantes.LEFT -> {
                        xssfCellStyle.setBorderBottom(getBorderStyle(value));
                        xssfCellStyle.setBottomBorderColor(xssfColor);
                    }
                    case CssConstantes.TOP -> {
                        xssfCellStyle.setBorderTop(getBorderStyle(value));
                        xssfCellStyle.setTopBorderColor(xssfColor);
                    }
                    case CssConstantes.BOTTOM -> {
                        xssfCellStyle.setBorderLeft(getBorderStyle(value));
                        xssfCellStyle.setLeftBorderColor(xssfColor);
                    }
                    default -> {
                        xssfCellStyle.setBorderRight(getBorderStyle(value));
                        xssfCellStyle.setRightBorderColor(xssfColor);
                    }
                }
            }
        });
    }

    private BorderStyle getBorderStyle(String css) {
        final String borderPropertiesWithoutColor = Arrays.stream(css.split(" "))
                .filter(e -> e.contains("1p") || e.contains("2p") || e.contains("sol") || e.contains("das"))
                .collect(joining(" "));
        return CssConstantes.CSS_BORDER_VALUE_TO_BORDER_STYLE_MAP.getOrDefault(borderPropertiesWithoutColor, NONE);

    }

    private XSSFColor getBorderColor(String css) {
        final XSSFColor xssfColor = new XSSFColor();
        final String[] borderStyleComposition = css.split(SPACE);
        final String borderColor = Arrays.stream(borderStyleComposition)
                .filter(e -> e.contains(CssConstantes.HASHBANG))
                .findFirst()
                .orElse(CssConstantes.NORMAL_GRAY_BORDER_COLOR_HEX);
        xssfColor.setARGBHex(borderColor.replace(CssConstantes.HASHBANG, EMPTY));
        return xssfColor;
    }
}
