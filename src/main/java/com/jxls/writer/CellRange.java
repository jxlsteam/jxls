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

    public CellRange(Cell startCell, int width, int height) {
        this.startCell = startCell;
        this.width = width;
        this.height = height;
        cells = new Cell[width][];
        for(int col = 0; col < width; col++){
            cells[col] = new Cell[height];
            for(int row = 0; row < height; row++){
                cells[col][row] = new Cell(col, row);
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
        for(int i = col+1; i < width; i++){
            for(int j = startRow; j <= endRow; j++){
                cells[i][j].setCol( cells[i][j].getCol() + colShift );
            }
        }
    }
    
    public void shiftCellsWithColBlock(int startCol, int endCol, int row, int rowShift){
        for(int i = startCol; i <= endCol; i++){
            for(int j = row + 1; j < height; j++){
                cells[i][j].setRow(cells[i][j].getRow() + rowShift );
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

}
