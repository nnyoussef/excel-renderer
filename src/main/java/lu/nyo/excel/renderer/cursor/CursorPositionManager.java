package lu.nyo.excel.renderer.cursor;

import org.apache.poi.ss.util.CellRangeAddress;

import java.util.LinkedList;

public final class CursorPositionManager {

    private static class CellRangeAddressNode {
        private CellRangeAddressNode next;
        private CellRangeAddressNode previous;
        private CellRangeAddress value;
    }

    final CursorPosition cursorPosition;

    //is used as starting point from where we will calculate cell positioning. The idea of it is to provide an optimal start point for scanning
    //it's the head of a linked list, with elements of higher rowspan that the current cursor because they may represent a late collision
    //it guarantees that the cursor can see the blocked cells
    CellRangeAddressNode cursorVisibleBlockedCells = null;

    final LinkedList<CellRangeAddress> newlyAddedCellsToCurrentSelectedRow = new LinkedList<>();

    public CursorPositionManager(CursorPosition cursorPosition) {
        this.cursorPosition = cursorPosition;
    }

    public void add(CellRangeAddress cellAddresses) {
        newlyAddedCellsToCurrentSelectedRow.add(cellAddresses);
        int colSpan = cellAddresses.getLastColumn() - cellAddresses.getFirstColumn() + 1;
        cursorPosition.incrementPosition(0, colSpan);
    }

    private void flushNewlyAddedCellsToBlockedCellsLinkedList() {

        for (CellRangeAddress cellAddresses : newlyAddedCellsToCurrentSelectedRow) {
            CellRangeAddressNode visited = cursorVisibleBlockedCells;
            CellRangeAddressNode hold = null;
            if (cursorVisibleBlockedCells == null) {
                cursorVisibleBlockedCells = new CellRangeAddressNode();
                cursorVisibleBlockedCells.value = cellAddresses;
                continue;
            }

            while (visited != null) {
                if (visited.value.getFirstColumn() > cellAddresses.getFirstColumn()) {
                    CellRangeAddressNode newNode = new CellRangeAddressNode();
                    newNode.value = cellAddresses;
                    newNode.next = visited;

                    if (visited.previous == null) {
                        visited.previous = newNode;
                        cursorVisibleBlockedCells = newNode;
                    } else {
                        CellRangeAddressNode previous = visited.previous;
                        newNode.previous = previous;
                        visited.previous = newNode;
                        previous.next = newNode;
                    }
                    break;
                }
                hold = visited;
                visited = visited.next;
            }

            if (visited == null) {
                visited = new CellRangeAddressNode();
                visited.value = cellAddresses;
                visited.previous = hold;
                hold.next = visited;
            }

        }
        newlyAddedCellsToCurrentSelectedRow.clear();
    }

    public void setCursorToNextUnblockedPositionOnNewRow() {
        flushNewlyAddedCellsToBlockedCellsLinkedList();
        CellRangeAddressNode scanned = cursorVisibleBlockedCells;
        int minRowPosition = cursorVisibleBlockedCells.value.getLastRow();
        int minColPosition = cursorVisibleBlockedCells.value.getLastColumn();
        while (scanned != null) {
            if (isLessThan(scanned.value.getLastRow(), scanned.value.getFirstColumn(), minRowPosition, minColPosition)) {
                minRowPosition = scanned.value.getLastRow();
                minColPosition = scanned.value.getFirstColumn();
            }
            scanned = scanned.next;
        }
        cursorPosition.setPosition(minRowPosition + 1, minColPosition);
        clean();
    }

    public void setCursorToNextUnmergedColOnCurrentRow() {
        CellRangeAddressNode scanned = cursorVisibleBlockedCells;
        int closestUnmergedCellForRowAt = cursorPosition.getCellPosition();
        while (scanned != null) {
            CellRangeAddress cellAddresses = scanned.value;
            final int firstColumn = cellAddresses.getFirstColumn();
            final int lastColumn = cellAddresses.getLastColumn();
            final int currentColumnIndex = closestUnmergedCellForRowAt;
            if ((currentColumnIndex >= firstColumn && currentColumnIndex <= lastColumn)) {
                closestUnmergedCellForRowAt = lastColumn + 1;
            }
            scanned = scanned.next;
        }
        cursorPosition.setCellPosition(closestUnmergedCellForRowAt);
    }

    private boolean isLessThan(int row, int col, int row1, int col1) {
        if (row < row1) return true;
        else if (row == row1) return col < col1;
        return false;
    }

    private CellRangeAddressNode removeNodeFromChain(CellRangeAddressNode toRemove) {
        if (toRemove.next != null) toRemove.next.previous = toRemove.previous;
        if (toRemove.previous != null) toRemove.previous.next = toRemove.next;
        else cursorVisibleBlockedCells = toRemove.next;
        return toRemove.next;
    }

    private void clean() {
        CellRangeAddressNode scanned = cursorVisibleBlockedCells;
        while (scanned != null) {
            if (cursorPosition.getRowPosition() > scanned.value.getLastRow()) {
                scanned = removeNodeFromChain(scanned);
                continue;
            }
            scanned = scanned.next;
        }
    }

}
