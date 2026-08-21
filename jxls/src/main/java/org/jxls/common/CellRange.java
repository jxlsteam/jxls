package org.jxls.common;

import org.jxls.common.cellshift.CellShiftStrategy;
import org.jxls.common.cellshift.InnerCellShiftStrategy;

/**
 * Represents an Excel cell range
 * 
 * Keeps track of cell references in the range and their transformation
 * 
 * @author Leonid Vysochyn
 */
public class CellRange {
    enum CellStatus { DEFAULT, CHANGED, TRANSFORMED}

    private CellShiftStrategy cellShiftStrategy = new InnerCellShiftStrategy();
    private int width;
    private int height;
    private CellRef[][] cells;  // map cells in the range to their transformed cell.
                                // There is some confusion between null and NONE cells
    private CellStatus[][] statusMatrix; // for each cell in range, its status
    private int[] rowWidths;  // number of columns in each row, considering transformations
    private int[] colHeights; // number of rows in each column, considering transformations
    private CellRef startCellRef; // top left cell of the range

    public CellRange(CellRef startCell, int width, int height) {
        this.startCellRef = startCell;
        String sheetName = startCell.getSheetName();
        this.width = width;
        this.height = height;
        cells = new CellRef[height][];
        statusMatrix = new CellStatus[height][];
        colHeights = new int[width];
        rowWidths = new int[height];
        for (int row = 0; row < height; row++) {
            rowWidths[row] = width;
            cells[row] = new CellRef[width];
            statusMatrix[row] = new CellStatus[width];
            for (int col = 0; col < width; col++) {
                cells[row][col] = new CellRef(sheetName, row, col);
            }
        }
        for (int col = 0; col < width; col++) {
            colHeights[col] = height;
        }
    }

    /**
     * Print cell range content to console, for debugging purposes.
     */
    void print() {
        for (int i = 0; i < height; i++) {
            for (int j = 0; j < width; j++) {
                String text = " ";
                if (!isEmpty(i, j)) {
                    CellRef absCellRef = new CellRef(
                            cells[i][j].getRow() + startCellRef.getRow(),
                            cells[i][j].getCol() + startCellRef.getCol());
                    text = absCellRef + " ";
                }
                System.out.print(text);
            }
            System.out.println();
        }
    }
    
    public CellRef getCell(int row, int col) {
        return cells[row][col];
    }

    public CellShiftStrategy getCellShiftStrategy() {
        return cellShiftStrategy;
    }

    public void setCellShiftStrategy(CellShiftStrategy cellShiftStrategy) {
        this.cellShiftStrategy = cellShiftStrategy;
    }

    void setCell(int row, int col, CellRef cellRef) {
        cells[row][col] = cellRef;
    }

    /**
     * Shift cells in the range after a columns of cells was moved to right.
     * For each cell in range, verify if move is required and if it is allowed (i.e. not already moved).
     * Mark cells as "changed".
     * Update cell coordinates of moved cells.
     *
     * @param startRow start row of the block
     * @param endRow end row of the block
     * @param col column of the block
     * @param colShift number of columns to shift
     * @param updateRowWidths whether to update rowWidths array
     */ 
    public void shiftCellsWithRowBlock(int startRow, int endRow, int col, int colShift, boolean updateRowWidths) {
        for (int i = 0; i < height; i++) {
            for (int j = 0; j < width; j++) {
                boolean requiresShifting = cellShiftStrategy.requiresColShifting(cells[i][j], startRow, endRow, col);
                if (requiresShifting && isHorizontalShiftAllowed(col, colShift, i, j)) {
                    cells[i][j].setCol(cells[i][j].getCol() + colShift);
                    statusMatrix[i][j] = CellStatus.CHANGED;
                }
            }
        }
        if (updateRowWidths) {
            int maxRow = Math.min(endRow, rowWidths.length - 1);
            for (int row = startRow; row <= maxRow; row++) {
                rowWidths[row] += colShift;
            }
        }
    }

    /**
     * Horizontal shift of a cell (cellRow, cellCol) is not allowed if cell is already marked as "changed",
     * or if widthChange < 0 and there is an empty between col and cellCol.
     */
    private boolean isHorizontalShiftAllowed(int col, int widthChange, int cellRow, int cellCol) {
        if (statusMatrix[cellRow][cellCol] == CellStatus.CHANGED) {
            return false;
        }
        if (widthChange >= 0) {
            return true;
        }
        for (int i = cellCol - 1; i > col; i--) {
            if (isEmpty(cellRow, i)) {
                return false;
            }
        }
        return true;
    }

    public boolean requiresColShifting(CellRef cell, int startRow, int endRow, int startColShift) {
        return cellShiftStrategy.requiresColShifting(cell, startRow, endRow, startColShift);
    }

    /**
     * Shift cells in the range after a row of cells was moved downwards.
     * For each cell in range, verify if downwards move is required and if it is allowed (i.e. not already moved).
     * Mark cells as "changed".
     * Update cell coordinates of moved cells.
     *
     * @param startCol start column of the block
     * @param endCol end column of the block
     * @param row row of the block
     * @param rowShift number of rows to shift
     * @param updateColHeights whether to update colHeights array
     */ 
    public void shiftCellsWithColBlock(int startCol, int endCol, int row, int rowShift, boolean updateColHeights) {
        for (int i = 0; i < height; i++) {
            for (int j = 0; j < width; j++) {
                boolean requiresShifting = cellShiftStrategy.requiresRowShifting(cells[i][j], startCol, endCol, row);
                if (requiresShifting && isVerticalShiftAllowed(row, rowShift, i, j)) {
                    cells[i][j].setRow(cells[i][j].getRow() + rowShift);
                    statusMatrix[i][j] = CellStatus.CHANGED;
                }
            }
        }
        if (updateColHeights) {
            int maxCol = Math.min(endCol, colHeights.length - 1);
            for (int col = startCol; col <= maxCol; col++) {
                colHeights[col] += rowShift;
            }
        }
    }

    /**
     * Vertical shift of a cell (cellRow, cellCol) is not allowed if cell is already marked as CHANGED,
     * or if heightChange < 0 and there is an empty cell between row and cellRow.
     */
    private boolean isVerticalShiftAllowed(int row, int heightChange, int cellRow, int cellCol) {
        if (statusMatrix[cellRow][cellCol] == CellStatus.CHANGED) {
            return false;
        }
        if (heightChange >= 0) {
            return true;
        }
        for (int i = cellRow - 1; i > row; i--) {
            if (isEmpty(i, cellCol)) {
                return false;
            }
        }
        return true;
    }

    /**
     * Mark given cells as excluded (i.e. set to null).
     * 
     * Notice that isExcluded() checks for both null and NONE
     */
    public void excludeCells(int startCol, int endCol, int startRow, int endRow) {
        for (int row = startRow; row <= endRow; row++) {
            for (int col = startCol; col <= endCol; col++) {
                cells[row][col] = null;
            }
        }
    }

    /**
     * Set given cells to CellRef.NONE.
     * 
     * Notice that isExcluded() checks for both null and NONE
     */
    public void clearCells(int startCol, int endCol, int startRow, int endRow) {
        for (int row = startRow; row <= endRow; row++) {
            for (int col = startCol; col <= endCol; col++) {
                cells[row][col] = CellRef.NONE;
            }
        }
    }

    /**
     * Calculate height (in cells) of the target cell range.
     * 
     * Return max of colHeights array
     */
    public int calculateHeight() {
        int maxHeight = 0;
        for (int col = 0; col < width; col++) {
            maxHeight = Math.max(maxHeight, colHeights[col]);
        }
        return maxHeight;
    }

    /**
     * Calculate width (in cells) of the target cell range.
     * 
     * Return max of rowWidths array
     */
    public int calculateWidth() {
        int maxWidth = 0;
        for (int row = 0; row < height; row++) {
            maxWidth = Math.max(maxWidth, rowWidths[row]);
        }
        return maxWidth;
    }

    /**
     * True if given cell is excluded (i.e. null or CellRef.NONE)
     */
    public boolean isExcluded(int row, int col) {
        return !contains(row, col) || cells[row][col] == null || CellRef.NONE.equals(cells[row][col]);
    }

    public void markAsTransformed(int row, int col) {
        statusMatrix[row][col] = CellStatus.TRANSFORMED;
    }

    /** True if given cell is marked as "transformed" */
    public boolean isTransformed(int row, int col) {
        return statusMatrix[row][col] == CellStatus.TRANSFORMED;
    }

    /**
     * True if given (relative) cell is contained in the original, not transformed, range. 
     */
    public boolean contains(int row, int col) {
        return row >= 0 && row < cells.length && col >= 0 && cells[0].length > col;
    }

    /**
     * True if given row contains any cells marked as "excluded".
     * Presence of excluded cells suggests the presence of a command in the row.
     */
    public boolean containsCommandsInRow(int row) {
        for (int col = 0; col < width; col++) {
            if (isExcluded(row, col)) {
                return true;
            }
        }
        return false;
    }

    public boolean isEmpty(int row, int col) {
        return cells[row][col] == null;
    }

    /** True if given cell is marked as "changed" */
    public boolean hasChanged(int row, int col) {
        return statusMatrix[row][col] == CellStatus.CHANGED;
    }

    /**
     * Reset cells from "changed" to "default" status.
     * Do not reset cells marked as "transformed".
     */
    public void resetChangeMatrix() {
        for (int i = 0; i < height; i++) {
            for (int j = 0; j < width; j++) {
                if (statusMatrix[i][j] != CellStatus.TRANSFORMED) {
                    statusMatrix[i][j] = CellStatus.DEFAULT;
                }
            }
        }
    }

    /**
     * Loop over all cells in given source row and find the maximum row number of the target cells.
     */
    public int findTargetRow(int srcRow) {
        int maxRow = -1;
        for (int col = 0; col < width; col++) {
            CellRef cellRef = cells[srcRow][col];
            maxRow = Math.max(maxRow, cellRef.getRow());
        }
        if (maxRow < 0) {
            maxRow = srcRow;
        }
        return maxRow;
    }
}
