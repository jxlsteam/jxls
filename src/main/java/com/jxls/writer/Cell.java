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

    public Cell(int row, int col) {
        this(0, row, col);
    }

    public Cell(int sheetIndex, int row, int col) {
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

    public void setSheetIndex(int sheetIndex) {
        this.sheetIndex = sheetIndex;
    }

    public void setRow(int row) {
        this.row = row;
    }

    public void setCol(int col) {
        this.col = col;
    }

    public Cell add(Size size){
        return new Cell(row + size.getHeight(), col + size.getWidth());
    }

    public Cell append(Size size){
        col += size.getWidth();
        row += size.getHeight();
        return this;
    }

    @Override
    public String toString() {
        return "Cell(" + sheetIndex + "," + row + "," + col + ")";
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;

        Cell cell = (Cell) o;

        if (col != cell.col) return false;
        if (row != cell.row) return false;
        if (sheetIndex != cell.sheetIndex) return false;

        return true;
    }

    @Override
    public int hashCode() {
        int result = sheetIndex;
        result = 31 * result + row;
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
        return new Cell(row + yChange, col);
    }

    public Cell addXModification(int xChange) {
        return new Cell(row, col + xChange);
    }
}
