package lu.nyo.excel.renderer;


import lu.nyo.excel.renderer.cursor.CursorPosition;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.Map;
import java.util.ServiceLoader;
import java.util.stream.Collectors;

import static java.util.function.Function.identity;

public final class ExcelFileRenderer {

    private ExcelFileRenderer() throws IllegalAccessException {
        throw new IllegalAccessException();
    }

    private static final Map<Object, RenderingEngine<?>> EXCEL_RENDERING_ENGINES = getExcelElementRendererLookup();

    private static Map getExcelElementRendererLookup() {
        return ServiceLoader.load(RenderingEngine.class)
                .stream()
                .map(ServiceLoader.Provider::get)
                .collect(Collectors.toUnmodifiableMap(RenderingEngine::getSelector, identity()));
    }

    public static void addPlugin() {
        throw new UnsupportedOperationException("Not implemented");
    }

    /**
     * @param css                        the css for styling
     * @param outputStream               the OutputStream or File to write on
     * @param rendereableObjectsPerSheet a map contain sheets content, the key is the title of
     *                                   the sheet, the value is a list of object to render on
     *                                   each sheet
     */
    public static void render(Map<String, LinkedList<Renderable>> rendereableObjectsPerSheet,
                              String css,
                              OutputStream outputStream) {

        try (SXSSFWorkbook workbook = new SXSSFWorkbook(100)) {
            final CellStyleProcessor cellStyleProcessor = CellStyleProcessor.create(css, workbook);
            workbook.setCompressTempFiles(true);
            rendereableObjectsPerSheet.forEach((key, value) -> {
                SXSSFSheet xssfSheet = workbook.createSheet(key);
                CursorPosition cursorPosition = new CursorPosition();
                Iterator<? extends Renderable> excelElementPerSheet = value.iterator();

                while (excelElementPerSheet.hasNext()) {
                    Renderable renderable = excelElementPerSheet.next();
                    Object selector = renderable.getRenderingEngineSelector();
                    RenderingEngine renderingEngine = EXCEL_RENDERING_ENGINES.get(selector);
                    renderingEngine.render(cursorPosition, renderable, xssfSheet, cellStyleProcessor);
                    excelElementPerSheet.remove();
                }
                flushRows(xssfSheet);
                flushBufferedData(xssfSheet);
            });
            workbook.write(outputStream);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void flushRows(SXSSFSheet sheet) {
        try {
            sheet.flushRows();
        } catch (IOException e) {
            //ignore
        }
    }

    private static void flushBufferedData(SXSSFSheet sxssfSheet) {
        try {
            sxssfSheet.flushBufferedData();
        } catch (IOException e) {
            //ignore
        }
    }
}
