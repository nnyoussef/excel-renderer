package lu.nyo.excel.renderer.cursor;

public final class CursorPosition {
    private int rowPosition = 1;

    private int cellPosition = 1;

    public int getRowPosition() {
        return rowPosition;
    }

    public int getCellPosition() {
        return cellPosition;
    }

    void setCellPosition(int cellPosition) {
        this.cellPosition = cellPosition;
    }

    public void incrementPosition(int rowIncrement,
                                  int colIncrement) {
        this.rowPosition += rowIncrement;
        this.cellPosition += colIncrement;
    }

    public void setPosition(int rowPosition,
                            int cellPosition) {
        this.rowPosition = rowPosition;
        this.cellPosition = cellPosition;
    }

}
