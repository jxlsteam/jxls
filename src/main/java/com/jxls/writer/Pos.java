package com.jxls.writer;

/**
 * @author Leonid Vysochyn
 *         Date: 1/25/12 1:42 PM
 */
public class Pos {
    int col;
    int row;
    int sheet;

    public Pos(int sheet, int row, int col) {
        this.sheet = sheet;
        this.row = row;
        this.col = col;
    }

    public Pos(int row, int col) {
        this(0, row, col);
    }
    
    public static Pos createFromCellRef(String cellRef){
        return new Pos(0,0,0);
    }
    
    public String getCellName(){
        return "A" + row;
    }

    public int getSheet() {
        return sheet;
    }

    public void setSheet(int sheet) {
        this.sheet = sheet;
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

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;

        Pos pos = (Pos) o;

        if (col != pos.col) return false;
        if (row != pos.row) return false;
        if (sheet != pos.sheet) return false;

        return true;
    }

    @Override
    public int hashCode() {
        int result = col;
        result = 31 * result + row;
        result = 31 * result + sheet;
        return result;
    }
}
