package lu.nyo.excel.renderer;

import com.steadystate.css.parser.CSSOMParser;
import com.steadystate.css.parser.SACParserCSS3;
import lu.nyo.excel.renderer.styleprocessor.BorderProcessor;
import lu.nyo.excel.renderer.styleprocessor.CellBackgroundProcessor;
import lu.nyo.excel.renderer.styleprocessor.TextAlignmentProcessor;
import lu.nyo.excel.renderer.styleprocessor.TextFontProcessor;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.w3c.css.sac.InputSource;
import org.w3c.dom.css.CSSRule;
import org.w3c.dom.css.CSSRuleList;
import org.w3c.dom.css.CSSStyleDeclaration;
import org.w3c.dom.css.CSSStyleSheet;

import java.io.IOException;
import java.io.StringReader;
import java.util.HashMap;
import java.util.Map;
import java.util.function.BiConsumer;

import static lu.nyo.excel.renderer.constantes.CssConstantes.DEFAULT_CSS_PROPERTIES;

public final class CellStyleProcessor {

    private final Map<String, XSSFCellStyle> cacheCellStyle = HashMap.newHashMap(30);
    private final SXSSFWorkbook xssfWorkbook;
    private final Map<String, Map<String, String>> cssRuleDeclaration;
    private final BiConsumer<Map<String, String>, XSSFCellStyle> styleProcessingChain;

    private CellStyleProcessor(SXSSFWorkbook xssfWorkbook,
                               Map<String, Map<String, String>> cssRuleDeclaration) {
        this.xssfWorkbook = xssfWorkbook;
        this.cssRuleDeclaration = cssRuleDeclaration;
        this.styleProcessingChain = new BorderProcessor()
                .andThen(new TextAlignmentProcessor())
                .andThen(new CellBackgroundProcessor())
                .andThen(new TextFontProcessor(() -> ((XSSFFont) xssfWorkbook.createFont())));
    }

    public static CellStyleProcessor create(String css,
                                            SXSSFWorkbook xssfWorkbook) throws IOException {
        final Map<String, Map<String, String>> cssRuleDeclaration = HashMap.newHashMap(30);
        final InputSource inputSource = new InputSource(new StringReader(css));
        final CSSOMParser parser = new CSSOMParser(new SACParserCSS3());
        final CSSStyleSheet styleSheet1 = parser.parseStyleSheet(inputSource, null, null);

        final CSSRuleList rules = styleSheet1.getCssRules();
        for (int i = 0; i < rules.getLength(); i++) {
            final CSSRule rule = rules.item(i);

            final String ruleName = rule.getCssText().split("[{]")[0].strip().trim();
            cssRuleDeclaration.put(ruleName, HashMap.newHashMap(12));

            String ruleDeclaration = rule.getCssText().replace("}", "");
            ruleDeclaration = ruleDeclaration.split("[{]")[1].trim().strip();

            final InputSource source = new InputSource(new StringReader(ruleDeclaration));
            final CSSStyleDeclaration declaration = parser.parseStyleDeclaration(source);

            for (int j = 0; j < declaration.getLength(); j++) {
                final String propName = declaration.item(j).trim().strip();
                final String propertyValue = declaration.getPropertyValue(propName);
                cssRuleDeclaration.get(ruleName).put(propName, propertyValue.trim().strip());
            }
        }
        return new CellStyleProcessor(xssfWorkbook, cssRuleDeclaration);
    }

    public XSSFCellStyle createStyle(String css) {
        css = css.trim().strip();

        if (cacheCellStyle.containsKey(css)) return cacheCellStyle.get(css);

        final XSSFCellStyle xssfCellStyle = (XSSFCellStyle) xssfWorkbook.createCellStyle();

        final Map<String, String> cssProperties = cssRuleDeclaration.getOrDefault(css, DEFAULT_CSS_PROPERTIES);

        styleProcessingChain.accept(cssProperties, xssfCellStyle);
        cacheCellStyle.put(css, xssfCellStyle);
        return xssfCellStyle;
    }

}