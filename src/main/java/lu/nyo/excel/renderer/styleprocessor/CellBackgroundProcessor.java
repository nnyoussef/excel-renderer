package lu.nyo.excel.renderer.styleprocessor;

import lu.nyo.excel.renderer.constantes.CssConstantes;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import java.util.Map;
import java.util.function.BiConsumer;

import static lu.nyo.excel.renderer.utils.ColorUtils.getXSSFColorFromRgb;
import static org.apache.poi.ss.usermodel.FillPatternType.SOLID_FOREGROUND;

public class CellBackgroundProcessor implements BiConsumer<Map<String, String>, XSSFCellStyle> {
    @Override
    public void accept(Map<String, String> cssInstructions, XSSFCellStyle xssfCellStyle) {
        final String cellBackground = cssInstructions.get(CssConstantes.BACKGROUND);
        xssfCellStyle.setFillForegroundColor(getXSSFColorFromRgb(cellBackground));
        xssfCellStyle.setFillPattern(SOLID_FOREGROUND);
    }

}
