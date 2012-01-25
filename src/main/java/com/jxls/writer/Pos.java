package com.jxls.writer;

/**
 * @author Leonid Vysochyn
 *         Date: 1/25/12 1:42 PM
 */
public class Pos {
    int col;
    int row;

    public Pos(int col, int row) {
        this.col = col;
        this.row = row;
    }

    public int getCol() {
        return col;
    }

    public void setCol(int col) {
        this.col = col;
    }

    public int getRow() {
        return row;
    }

    public void setRow(int row) {
        this.row = row;
    }
}
