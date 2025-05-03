package lu.nyo.excel.renderer.styleprocessor;

import lu.nyo.excel.renderer.constantes.CssConstantes;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import java.util.Map;
import java.util.function.BiConsumer;

import static java.util.Optional.ofNullable;
import static org.apache.commons.lang3.StringUtils.EMPTY;
import static org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT;
import static org.apache.poi.ss.usermodel.HorizontalAlignment.RIGHT;
import static org.apache.poi.ss.usermodel.VerticalAlignment.CENTER;

public class TextAlignmentProcessor implements BiConsumer<Map<String, String>, XSSFCellStyle> {
    @Override
    public void accept(Map<String, String> cssInstruction, XSSFCellStyle xssfCellStyle) {
        final String textAlignment = cssInstruction.get(CssConstantes.TEXT_ALIGN);
        xssfCellStyle.setVerticalAlignment(CENTER);
        xssfCellStyle.setAlignment(getHorizontalAlignment(textAlignment));
    }

    private HorizontalAlignment getHorizontalAlignment(String textAlignment) {
        return switch (ofNullable(textAlignment).orElse(EMPTY)) {
            case CssConstantes.TEXT_ALIGN_CENTER -> HorizontalAlignment.CENTER;
            case CssConstantes.TEXT_ALIGN_RIGHT -> RIGHT;
            default -> LEFT;
        };
    }
}
