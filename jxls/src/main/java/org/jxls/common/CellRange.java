package org.jxls.common;

import org.jxls.common.cellshift.CellShiftStrategy;
import org.jxls.common.cellshift.InnerCellShiftStrategy;

/**
 * Represents an Excel cell range
 * 
 * @author Leonid Vysochyn
 */
public class CellRange {
    private CellShiftStrategy cellShiftStrategy = new InnerCellShiftStrategy();
    private int width;
    private int height;
    private CellRef[][] cells;
    private boolean[][] changeMatrix;
    private int[] rowWidths;
    private int[] colHeights;
    private CellRef startCellRef;

    public CellRange(CellRef startCell, int width, int height) {
        this.startCellRef = startCell;
        String sheetName = startCell.getSheetName();
        this.width = width;
        this.height = height;
        cells = new CellRef[height][];
        changeMatrix = new boolean[height][];
        colHeights = new int[width];
        rowWidths = new int[height];
        for (int row = 0; row < height; row++) {
            rowWidths[row] = width;
            cells[row] = new CellRef[width];
            changeMatrix[row] = new boolean[width];
            for (int col = 0; col < width; col++) {
                cells[row][col] = new CellRef(sheetName, row, col);
            }
        }
        for (int col = 0; col < width; col++) {
            colHeights[col] = height;
        }
    }

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

    public void shiftCellsWithRowBlock(int startRow, int endRow, int col, int colShift, boolean updateRowWidths) {
        for (int i = 0; i < height; i++) {
            for (int j = 0; j < width; j++) {
                boolean requiresShifting = cellShiftStrategy.requiresColShifting(cells[i][j], startRow, endRow, col);
                if (requiresShifting && isHorizontalShiftAllowed(col, colShift, i, j)) {
                    cells[i][j].setCol(cells[i][j].getCol() + colShift);
                    changeMatrix[i][j] = true;
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

    private boolean isHorizontalShiftAllowed(int col, int widthChange, int cellRow, int cellCol) {
        if (changeMatrix[cellRow][cellCol]) {
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

    public void shiftCellsWithColBlock(int startCol, int endCol, int row, int rowShift, boolean updateColHeights) {
        for (int i = 0; i < height; i++) {
            for (int j = 0; j < width; j++) {
                boolean requiresShifting = cellShiftStrategy.requiresRowShifting(cells[i][j], startCol, endCol, row);
                if (requiresShifting && isVerticalShiftAllowed(row, rowShift, i, j)) {
                    cells[i][j].setRow(cells[i][j].getRow() + rowShift);
                    changeMatrix[i][j] = true;
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

    private boolean isVerticalShiftAllowed(int row, int heightChange, int cellRow, int cellCol) {
        if (changeMatrix[cellRow][cellCol]) {
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

    public void excludeCells(int startCol, int endCol, int startRow, int endRow) {
        for (int row = startRow; row <= endRow; row++) {
            for (int col = startCol; col <= endCol; col++) {
                cells[row][col] = null;
            }
        }
    }

    public void clearCells(int startCol, int endCol, int startRow, int endRow) {
        for (int row = startRow; row <= endRow; row++) {
            for (int col = startCol; col <= endCol; col++) {
                cells[row][col] = CellRef.NONE;
            }
        }
    }

    public int calculateHeight() {
        int maxHeight = 0;
        for (int col = 0; col < width; col++) {
            maxHeight = Math.max(maxHeight, colHeights[col]);
        }
        return maxHeight;
    }

    public int calculateWidth() {
        int maxWidth = 0;
        for (int row = 0; row < height; row++) {
            maxWidth = Math.max(maxWidth, rowWidths[row]);
        }
        return maxWidth;
    }

    public boolean isExcluded(int row, int col) {
        return cells[row][col] == null || CellRef.NONE.equals(cells[row][col]);
    }

    public boolean contains(int row, int col){
        return row >= 0 && row < cells.length && col >= 0 && cells[0].length > col;
    }

    public boolean containsCommandsInRow(int row) {
        for (int col = 0; col < width; col++) {
            if (cells[row][col] == null || cells[row][col] == CellRef.NONE) {
                return true;
            }
        }
        return false;
    }

    public boolean isEmpty(int row, int col) {
        return cells[row][col] == null;
    }

    public boolean hasChanged(int row, int col) {
        return changeMatrix[row][col];
    }

    public void resetChangeMatrix() {
        for (int i = 0; i < height; i++) {
            for (int j = 0; j < width; j++) {
                changeMatrix[i][j] = false;
            }
        }
    }

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
