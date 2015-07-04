package org.jxls.common;

/**
 * Represents an excel cell range
 * @author Leonid Vysochyn
 *         Date: 1/26/12
 */
public class CellRange {
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

    void setCell(int row, int col, CellRef cellRef){
        cells[row][col] = cellRef;
    }

    public void shiftCellsWithRowBlock(int startRow, int endRow, int col, int colShift){
        for(int i = 0; i < height; i++){
            for(int j = 0; j < width; j++){
                if( cells[i][j] != null && cells[i][j].getCol() > col && cells[i][j].getRow() >= startRow && cells[i][j].getRow() <= endRow && !changeMatrix[i][j]){
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
    
    public void shiftCellsWithColBlock(int startCol, int endCol, int row, int rowShift){
        for(int i = 0; i < height; i++){
            for(int j = 0; j < width; j++){
                if(cells[i][j] != null && cells[i][j].getRow() > row && cells[i][j].getCol() >= startCol && cells[i][j].getCol() <= endCol && !changeMatrix[i][j]){
                    cells[i][j].setRow(cells[i][j].getRow() + rowShift );
                    changeMatrix[i][j] = true;
                }
            }
        }
        int maxCol = Math.min(endCol, colHeights.length-1);
        for(int col = startCol; col <= maxCol; col++){
            colHeights[col] += rowShift;
        }
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
