package org.jxls.common;

import org.jxls.common.cellshift.CellShiftStrategy;
import org.jxls.common.cellshift.InnerCellShiftStrategy;

/**
 * Represents an excel cell range
 * @author Leonid Vysochyn
 *         Date: 1/26/12
 */
public class CellRange {
    private CellShiftStrategy cellShiftStrategy = new InnerCellShiftStrategy();
    CellRef startCell;
    int width;
    int height;
    CellRef[][] cells;
    boolean[][] changeMatrix;
    int[] rowWidths;
    int[] colHeights;

    public CellRange(CellRef startCell, int width, int height) {
        String sheetName = startCell.getSheetName();
        this.startCell = startCell;
        this.width = width;
        this.height = height;
        cells = new CellRef[height][];
        changeMatrix = new boolean[height][];
        colHeights = new int[width];
        rowWidths = new int[height];
        for(int row = 0; row < height; row++){
            rowWidths[row] = width;
            cells[row] = new CellRef[width];
            changeMatrix[row] = new boolean[width];
            for(int col = 0; col < width; col++){
                cells[row][col] = new CellRef(sheetName, row, col);
            }
        }
        for(int col = 0; col < width; col++){
            colHeights[col] = height;
        }
    }
    
    public CellRef getCell(int row, int col){
        return cells[row][col];
    }

    public CellShiftStrategy getCellShiftStrategy() {
        return cellShiftStrategy;
    }

    public void setCellShiftStrategy(CellShiftStrategy cellShiftStrategy) {
        this.cellShiftStrategy = cellShiftStrategy;
    }

    void setCell(int row, int col, CellRef cellRef){
        cells[row][col] = cellRef;
    }

    public void shiftCellsWithRowBlock(int startRow, int endRow, int col, int colShift, boolean updateRowWidths){
        for(int i = 0; i < height; i++){
            for(int j = 0; j < width; j++){
                boolean requiresShifting = cellShiftStrategy.requiresColShifting(cells[i][j], startRow, endRow, col);
                if(requiresShifting && !changeMatrix[i][j]){
                    cells[i][j].setCol(cells[i][j].getCol() + colShift);
                    changeMatrix[i][j] = true;
                }
            }
        }
        int maxRow = Math.min(endRow, rowWidths.length-1);
        for(int row = startRow; row <= maxRow; row++){
            rowWidths[row] += colShift;
        }
    }

    public boolean requiresColShifting(CellRef cell, int startRow, int endRow, int startColShift){
        return cellShiftStrategy.requiresColShifting(cell, startRow, endRow, startColShift);
    }

    public void shiftCellsWithColBlock(int startCol, int endCol, int row, int rowShift, boolean updateColHeights){
        for(int i = 0; i < height; i++){
            for(int j = 0; j < width; j++){
                boolean requiresShifting = cellShiftStrategy.requiresRowShifting(cells[i][j], startCol, endCol, row);
                if(requiresShifting && !changeMatrix[i][j]){
                    cells[i][j].setRow(cells[i][j].getRow() + rowShift );
                    changeMatrix[i][j] = true;
                }
            }
        }
        if( updateColHeights) {
            int maxCol = Math.min(endCol, colHeights.length - 1);
            for (int col = startCol; col <= maxCol; col++) {
                colHeights[col] += rowShift;
            }
        }
    }

    public boolean requiresRowShifting(CellRef cell, int startCol, int endCol, int startRowShift){
        return cellShiftStrategy.requiresRowShifting(cell, startCol, endCol, startRowShift);
    }

    public CellRef getStartCell() {
        return startCell;
    }
    
    public void excludeCells(int startCol, int endCol, int startRow, int endRow){
        for(int row = startRow; row <= endRow; row++){
            for(int col = startCol; col <= endCol; col++){
                cells[row][col] = null;
            }
        }
    }

    public int calculateHeight(){
        int maxHeight = 0;
        for(int col = 0; col < width; col++){
            maxHeight = Math.max(maxHeight, colHeights[col]);
        }
        return maxHeight;
    }

    public int calculateWidth(){
        int maxWidth = 0;
        for(int row = 0; row < height; row++){
            maxWidth = Math.max(maxWidth, rowWidths[row]);
        }
        return maxWidth;
    }

    public boolean isExcluded(int row, int col){
        return cells[row][col] == null;
    }

    public boolean hasChanged(int row, int col){
        return changeMatrix[row][col];
    }
    
    public void resetChangeMatrix(){
        for(int i = 0; i < height; i++){
            for(int j=0; j < width; j++){
                changeMatrix[i][j] = false;
            }
        }
    }

}
