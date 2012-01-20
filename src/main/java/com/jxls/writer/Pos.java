package com.jxls.writer;

/**
 * Date: Mar 13, 2009
 *
 * @author Leonid Vysochyn
 */
public class Pos {
    int col;
    int row;

    public Pos(int row, int col) {
        this.row = row;
        this.col = col;
    }

    public int getCol() {
        return col;
    }

    public int getRow() {
        return row;
    }

    public Pos add(Size size){
        return new Pos(row + size.getHeight(), col + size.getWidth());
    }

    public Pos append(Size size){
        row += size.getHeight();
        col += size.getWidth();
        return this;
    }

    @Override
    public String toString() {
        return "(" + row + "," + col + ")";    
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;

        Pos pos = (Pos) o;

        if (col != pos.col) return false;
        if (row != pos.row) return false;

        return true;
    }

    @Override
    public int hashCode() {
        int result = col;
        result = 31 * result + row;
        return result;
    }

    public boolean isLower(Pos pos) {
        return row > pos.getRow();
    }

    public boolean isOnTheRightOf(Pos pos){
        return col > pos.getCol();
    }

    public Size minus(Pos pos) {
        return new Size(col - pos.getCol(), row - pos.getRow());
    }

    public Pos addColModification(int colChange) {
        return new Pos(row, col + colChange);
    }

    public Pos addRowModification(int rowChange) {
        return new Pos(row + rowChange, col);
    }
}
