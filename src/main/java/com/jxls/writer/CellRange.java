package com.jxls.writer;

/**
 * @author Leonid Vysochyn
 *         Date: 1/26/12 1:43 PM
 */
public class CellRange {
    Cell startCell;
    int width;
    int height;
    Cell[][] cells;
    boolean[][] changeMatrix;

    public CellRange(Cell startCell, int width, int height) {
        int sheetIndex = startCell.getSheetIndex();
        this.startCell = startCell;
        this.width = width;
        this.height = height;
        cells = new Cell[width][];
        changeMatrix = new boolean[width][];
        for(int col = 0; col < width; col++){
            cells[col] = new Cell[height];
            changeMatrix[col] = new boolean[height];
            for(int row = 0; row < height; row++){
                cells[col][row] = new Cell(col, row, sheetIndex);
            }
        }
    }
    
    public Cell getCell(int col, int row){
        return cells[col][row];
    }

    void setCell(int col, int row, Cell cell){
        cells[col][row] = cell;
    }

    public void shiftCellsWithRowBlock(int startRow, int endRow, int col, int colShift){
        for(int i = 0; i < width; i++){
            for(int j = 0; j < height; j++){
                if( cells[i][j] != null && cells[i][j].getCol() > col && cells[i][j].getRow() >= startRow && cells[i][j].getRow() <= endRow && !changeMatrix[i][j]){
                    cells[i][j].setCol( cells[i][j].getCol() + colShift );
                    changeMatrix[i][j] = true;
                }
            }
        }
    }
    
    public void shiftCellsWithColBlock(int startCol, int endCol, int row, int rowShift){
        for(int i = 0; i < width; i++){
            for(int j = 0; j < height; j++){
                if(cells[i][j] != null && cells[i][j].getRow() > row && cells[i][j].getCol() >= startCol && cells[i][j].getCol() <= endCol && !changeMatrix[i][j]){
                    cells[i][j].setRow(cells[i][j].getRow() + rowShift );
                    changeMatrix[i][j] = true;
                }
            }
        }
    }
    
    public void excludeCells(int startCol, int endCol, int startRow, int endRow){
        for(int col = startCol; col <= endCol; col++){
            for(int row = startRow; row <= endRow; row++){
                cells[col][row] = null;
            }
        }
    }

    public boolean isExcluded(int col, int row){
        return cells[col][row] == null;
    }

    public boolean hasChanged(int col, int row){
        return changeMatrix[col][row];
    }
    
    public void resetChangeMatrix(){
        for(int i = 0; i < width; i++){
            for(int j=0; j<height; j++){
                changeMatrix[i][j] = false;
            }
        }
    }

}
