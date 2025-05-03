package lu.nyo.excel.renderer.styleprocessor;

import lu.nyo.excel.renderer.constantes.CssConstantes;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.util.Map;
import java.util.function.BiConsumer;
import java.util.function.Supplier;

import static java.util.Optional.ofNullable;
import static lu.nyo.excel.renderer.utils.ColorUtils.getXSSFColorFromRgb;

public final class TextFontProcessor implements BiConsumer<Map<String, String>, XSSFCellStyle> {

    private final Supplier<XSSFFont> xssfFontSupplier;

    public TextFontProcessor(Supplier<XSSFFont> xssfFontSupplier) {
        this.xssfFontSupplier = xssfFontSupplier;
    }

    @Override
    public void accept(Map<String, String> instructions,
                       XSSFCellStyle xssfCellStyle) {
        final XSSFColor textColor = getXSSFColorFromRgb(instructions.get(CssConstantes.COLOR));
        boolean isBold = isBold(instructions);

        final XSSFFont xssfFont = xssfFontSupplier.get();
        xssfFont.setColor(textColor);
        xssfFont.setBold(isBold);
        xssfCellStyle.setFont(xssfFont);
    }

    private boolean isBold(Map<String, String> cssProperties) {
        return ofNullable(cssProperties.get(CssConstantes.FONT_WEIGHT))
                .map(fw -> fw.contains(CssConstantes.BOLD))
                .orElse(false);
    }
}
