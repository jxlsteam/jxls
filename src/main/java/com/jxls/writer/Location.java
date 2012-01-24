package com.jxls.writer;

import java.util.List;
import java.util.ArrayList;

/**
 * Date: Mar 31, 2009
 *
 * @author Leonid Vysochyn
 */
public class Location {
    Cell cell;
    Size size;

    List<Cell> cellList;
    private XlsProxy xlsProxy;

    public Location(Cell cell, Size size) {
        this.cell = cell;
        this.size = size;
    }

    public Cell getCell() {
        return cell;
    }

    public Size getSize() {
        return size;
    }

    public int getRightmostCol(){
        return cell.getRow() + size.getWidth();
    }

    public int getLowestRow(){
        return cell.getCol() + size.getHeight();
    }

    List<Cell> getCellList(){
        if( cellList == null ){
            cellList = new ArrayList<Cell>();
            for( int row = cell.getCol(); row <= cell.getCol() + size.getHeight(); row++){
                for( int col = cell.getRow(); col <= cell.getRow() + size.getWidth(); col++){
                    cellList.add( new Cell(row, col));
                }
            }
        }
        return cellList;
    }

    @Override
    public String toString() {
        return "Cell: " + cell + ", Size: " + size;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;

        Location location = (Location) o;

        if (cell != null ? !cell.equals(location.cell) : location.cell != null) return false;
        if (size != null ? !size.equals(location.size) : location.size != null) return false;

        return true;
    }

    @Override
    public int hashCode() {
        int result = cell != null ? cell.hashCode() : 0;
        result = 31 * result + (size != null ? size.hashCode() : 0);
        return result;
    }

    public boolean isOnTheLeftOf(Location location) {
        if( cell.getRow() < location.getCell().getRow() ){
            int startRow = cell.getCol();
            int endRow = cell.getCol() + size.getHeight();
            int startRow2 = location.getCell().getCol();
            int endRow2 = location.getCell().getCol() + size.getHeight();
            return (startRow >= startRow2 && startRow <= endRow2 ) ||
                    (startRow2 >= startRow && startRow2 <= endRow);
        }else{
            return false;
        }
    }

    public boolean isAbove(Location location) {
        if( cell.getCol() < location.getCell().getCol() ){
            int startCol = cell.getRow();
            int endCol = cell.getRow() + size.getWidth();
            int startCol2 = location.getCell().getRow();
            int endCol2 = location.getCell().getRow() + location.getSize().getWidth();
            return (startCol >= startCol2 && startCol <= endCol2) ||
                    (startCol2 >= startCol && startCol2 <= endCol );
        }else{
            return false;
        }
    }

    public void setXlsProxy(XlsProxy xlsProxy) {
        this.xlsProxy = xlsProxy;
    }

    public boolean contains(Cell cell) {
        return containsRow(cell.getCol()) && containsCol(cell.getRow());
    }

    private boolean containsCol(int col) {
        return cell.getRow() <= col && cell.getRow() +
                size.getWidth() >= col;
    }

    private boolean containsRow(int row) {
        return cell.getCol() <= row && cell.getCol() +
                size.getHeight() >= row;
    }

    public boolean isOnTheLeftOf(Cell aCell) {
        if( cell.getRow() < aCell.getRow() ){
            int startRow = cell.getCol();
            int endRow = cell.getCol() + size.getHeight();
            return (startRow <= aCell.getCol() && endRow >= aCell.getCol() );
        }else{
            return false;
        }
    }

    public boolean isAbove(Cell aCell) {
        if( cell.getCol() < aCell.getCol() ){
            int startCol = cell.getRow();
            int endCol = cell.getRow() + size.getWidth();
            return (startCol <= aCell.getRow() && endCol >= aCell.getRow());
        }else{
            return false;
        }
    }

    public Location addPosColModification(int colChange) {
        Cell newCell = cell.addYModification(colChange);
        return new Location(newCell, size);
    }

    public Location addPosRowModification(int rowChange) {
        Cell newCell = cell.addXModification(rowChange);
        return new Location(newCell, size);
    }

    public Cell getLeftBottomPos() {
        return new Cell(cell.getCol() + size.getHeight(), cell.row);
    }
}
