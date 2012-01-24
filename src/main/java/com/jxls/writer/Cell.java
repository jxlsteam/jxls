package com.jxls.writer;

/**
 * Date: Mar 13, 2009
 *
 * @author Leonid Vysochyn
 */
public class Cell {
    int sheetIndex;
    int row;
    int col;

    public Cell(int col, int row) {
        this(col, row, 0);
    }

    public Cell(int col, int row, int sheetIndex) {
        this.col = col;
        this.row = row;
        this.sheetIndex = sheetIndex;
    }

    public int getRow() {
        return row;
    }

    public int getCol() {
        return col;
    }

    public int getSheetIndex() {
        return sheetIndex;
    }

    public Cell add(Size size){
        return new Cell(col + size.getWidth(), row + size.getHeight());
    }

    public Cell append(Size size){
        col += size.getWidth();
        row += size.getHeight();
        return this;
    }

    @Override
    public String toString() {
        return "(" + col + "," + row + ")";
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;

        Cell cell = (Cell) o;

        if (row != cell.row) return false;
        if (col != cell.col) return false;

        return true;
    }

    @Override
    public int hashCode() {
        int result = row;
        result = 31 * result + col;
        return result;
    }

    public boolean isLower(Cell cell) {
        return row > cell.getRow();
    }

    public boolean isOnTheRightOf(Cell cell){
        return col > cell.getCol();
    }

    public Size minus(Cell cell) {
        return new Size(col - cell.getCol(), row - cell.getRow());
    }

    public Cell addYModification(int yChange) {
        return new Cell(col, row + yChange);
    }

    public Cell addXModification(int xChange) {
        return new Cell(col + xChange, row);
    }
}
