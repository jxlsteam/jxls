package com.jxls.writer.common;

/**
 * @author Leonid Vysochyn
 *         Date: 1/26/12 1:43 PM
 */
public class CellRange {
    CellRef startCell;
    int width;
    int height;
    CellRef[][] cells;
    boolean[][] changeMatrix;

    public CellRange(CellRef startCell, int width, int height) {
        String sheetName = startCell.getSheetName();
        this.startCell = startCell;
        this.width = width;
        this.height = height;
        cells = new CellRef[height][];
        changeMatrix = new boolean[height][];
        for(int row = 0; row < height; row++){
            cells[row] = new CellRef[width];
            changeMatrix[row] = new boolean[width];
            for(int col = 0; col < width; col++){
                cells[row][col] = new CellRef(sheetName, row, col);
            }
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
                    cells[i][j].setCol( cells[i][j].getCol() + colShift );
                    changeMatrix[i][j] = true;
                }
            }
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
    }
    
    public void excludeCells(int startCol, int endCol, int startRow, int endRow){
        for(int row = startRow; row <= endRow; row++){
            for(int col = startCol; col <= endCol; col++){
                cells[row][col] = null;
            }
        }
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
