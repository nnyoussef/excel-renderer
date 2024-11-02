package lu.nyo.excel.renderer.engine;


import lu.nyo.excel.renderer.CellStyleProcessor;
import lu.nyo.excel.renderer.RenderingEngine;
import lu.nyo.excel.renderer.cursor.CursorPosition;
import lu.nyo.excel.renderer.cursor.CursorPositionManager;
import lu.nyo.excel.renderer.excel_element.Row;
import lu.nyo.excel.renderer.excel_element.Table;
import lu.nyo.excel.renderer.utils.CellDataUtils;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import java.util.HashMap;
import java.util.Map;
import java.util.concurrent.atomic.AtomicReference;
import java.util.stream.Stream;

import static lu.nyo.excel.renderer.utils.SpanUtils.createSpan;

public class TableRenderingEngine implements RenderingEngine<Table> {
    private static final Map<String, String> CELL_CSS_FULL_CLASS_NAMES_CACHE_TBODY = new HashMap<>(50);

    static {
        CELL_CSS_FULL_CLASS_NAMES_CACHE_TBODY.put("", "tbody .default");
    }

    @Override
    public void handle(CursorPosition cursorPosition,
                       Table table,
                       SXSSFSheet worksheet,
                       CellStyleProcessor cellStyleProcessor) {

        AtomicReference<Stream<Row>> header = new AtomicReference<>(Stream.empty());
        AtomicReference<Stream<Row>> body = new AtomicReference<>(Stream.empty());
        AtomicReference<Stream<Row>> footer = new AtomicReference<>(Stream.empty());

        table.tableContentInitializer(header, body, footer);

        render(worksheet, header.get(), cursorPosition, "thead", cellStyleProcessor);
        render(worksheet, body.get(), cursorPosition, "tbody", cellStyleProcessor);
        render(worksheet, footer.get(), cursorPosition, "tfoot", cellStyleProcessor);

        cursorPosition.resetCellPositionOnNextLine();
    }

    @Override
    public Class<Table> getSelector() {
        return Table.class;
    }

    private void render(final SXSSFSheet worksheet,
                        final Stream<Row> rows,
                        final CursorPosition cursorPosition,
                        final String tableSection,
                        final CellStyleProcessor cellStyleProcessor) {

        final CursorPositionManager cursorPositionManager = new CursorPositionManager(cursorPosition);

        rows.forEach(row -> {

            row.getCells().forEach(cell -> {
                final int rowSpan = cell.getRowSpan();
                final int colSpan = cell.getColSpan();
                final String css = getCssClassFullName(tableSection, cell.getCssClass());

                cursorPositionManager.setCursorToNextAvailableUnmergedColOnCurrentRow();

                final SXSSFRow xssfRow = (SXSSFRow) CellUtil.getRow(cursorPosition.getRowPosition(), worksheet);
                final SXSSFCell xssfCell = xssfRow.createCell(cursorPosition.getCellPosition());
                final XSSFCellStyle xssfCellStyle = cellStyleProcessor.createStyle(css);
                final CellRangeAddress cellAddressesAfterMerging = createSpan(worksheet, rowSpan, colSpan, cursorPosition);

                cursorPositionManager.add(cellAddressesAfterMerging);

                CellDataUtils.setData(cell.getData(), xssfCell);
                resolveStyleForMergedCell(cellAddressesAfterMerging, xssfCellStyle, worksheet);
                xssfCell.setCellStyle(xssfCellStyle);
            });
            cursorPositionManager.setCursorToNextAvailablePosition();
        });
    }

    private String getCssClassFullName(String tableSection,
                                       String cellCss) {
        String css;
        if ("tbody" == tableSection) {
            if (CELL_CSS_FULL_CLASS_NAMES_CACHE_TBODY.containsKey(cellCss)) {
                css = CELL_CSS_FULL_CLASS_NAMES_CACHE_TBODY.get(cellCss);
            } else {
                css = tableSection + " ." + cellCss;
                CELL_CSS_FULL_CLASS_NAMES_CACHE_TBODY.put(cellCss, css);
            }
        } else {
            if ("default" == cellCss)
                css = tableSection + " .default";
            else {
                css = tableSection + " ." + cellCss;
            }
        }
        return css;
    }

    private void resolveStyleForMergedCell(CellRangeAddress cellRangeAddress,
                                           XSSFCellStyle xssfCellStyle,
                                           SXSSFSheet sheet) {
        if (cellRangeAddress.getNumberOfCells() == 1) return;

        SXSSFRow lastRow = (SXSSFRow) CellUtil.getRow(cellRangeAddress.getLastRow(), sheet);

        for (int i = cellRangeAddress.getFirstColumn(); i < cellRangeAddress.getLastColumn() + 1; i++) {
            final SXSSFCell xssfCell = (SXSSFCell) CellUtil.getCell(lastRow, i);
            xssfCell.setCellStyle(xssfCellStyle);
        }

        for (int i = cellRangeAddress.getFirstRow(); i < cellRangeAddress.getLastRow() + 1; i++) {
            SXSSFRow row = (SXSSFRow) CellUtil.getRow(i, sheet);
            final SXSSFCell xssfCell = (SXSSFCell) CellUtil.getCell(row, cellRangeAddress.getLastColumn());
            xssfCell.setCellStyle(xssfCellStyle);
        }
    }
}
