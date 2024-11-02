package lu.nyo.excel.renderer;

import lu.nyo.excel.renderer.cursor.CursorPosition;
import org.apache.poi.xssf.streaming.SXSSFSheet;

public interface RenderingEngine<T> {

    void handle(CursorPosition cursorPosition,
                T elementToHandle,
                SXSSFSheet worksheet,
                CellStyleProcessor cellStyleProcessor);

    Class<T> getSelector();
}
