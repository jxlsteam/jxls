package com.jxls.writer;

import java.util.List;
import java.util.ArrayList;

/**
 * Date: Mar 31, 2009
 *
 * @author Leonid Vysochyn
 */
public class Location {
    Pos pos;
    Size size;

    List<Pos> posList;
    private XlsProxy xlsProxy;

    public Location(Pos pos, Size size) {
        this.pos = pos;
        this.size = size;
    }

    public Pos getPos() {
        return pos;
    }

    public Size getSize() {
        return size;
    }

    public int getRightmostCol(){
        return pos.getY() + size.getWidth();
    }

    public int getLowestRow(){
        return pos.getX() + size.getHeight();
    }

    List<Pos> getPosList(){
        if( posList == null ){
            posList = new ArrayList<Pos>();
            for( int row = pos.getX(); row <= pos.getX() + size.getHeight(); row++){
                for( int col = pos.getY(); col <= pos.getY() + size.getWidth(); col++){
                    posList.add( new Pos(row, col));
                }
            }
        }
        return posList;
    }

    @Override
    public String toString() {
        return "Pos: " + pos + ", Size: " + size;    
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;

        Location location = (Location) o;

        if (pos != null ? !pos.equals(location.pos) : location.pos != null) return false;
        if (size != null ? !size.equals(location.size) : location.size != null) return false;

        return true;
    }

    @Override
    public int hashCode() {
        int result = pos != null ? pos.hashCode() : 0;
        result = 31 * result + (size != null ? size.hashCode() : 0);
        return result;
    }

    public boolean isOnTheLeftOf(Location location) {
        if( pos.getY() < location.getPos().getY() ){
            int startRow = pos.getX();
            int endRow = pos.getX() + size.getHeight();
            int startRow2 = location.getPos().getX();
            int endRow2 = location.getPos().getX() + size.getHeight();
            return (startRow >= startRow2 && startRow <= endRow2 ) ||
                    (startRow2 >= startRow && startRow2 <= endRow);
        }else{
            return false;
        }
    }

    public boolean isAbove(Location location) {
        if( pos.getX() < location.getPos().getX() ){
            int startCol = pos.getY();
            int endCol = pos.getY() + size.getWidth();
            int startCol2 = location.getPos().getY();
            int endCol2 = location.getPos().getY() + location.getSize().getWidth();
            return (startCol >= startCol2 && startCol <= endCol2) ||
                    (startCol2 >= startCol && startCol2 <= endCol );
        }else{
            return false;
        }
    }

    public void setXlsProxy(XlsProxy xlsProxy) {
        this.xlsProxy = xlsProxy;
    }

    public boolean contains(Pos pos) {
        return containsRow(pos.getX()) && containsCol(pos.getY());
    }

    private boolean containsCol(int col) {
        return pos.getY() <= col && pos.getY() +
                size.getWidth() >= col;
    }

    private boolean containsRow(int row) {
        return pos.getX() <= row && pos.getX() +
                size.getHeight() >= row;
    }

    public boolean isOnTheLeftOf(Pos aPos) {
        if( pos.getY() < aPos.getY() ){
            int startRow = pos.getX();
            int endRow = pos.getX() + size.getHeight();
            return (startRow <= aPos.getX() && endRow >= aPos.getX() );
        }else{
            return false;
        }
    }

    public boolean isAbove(Pos aPos) {
        if( pos.getX() < aPos.getX() ){
            int startCol = pos.getY();
            int endCol = pos.getY() + size.getWidth();
            return (startCol <= aPos.getY() && endCol >= aPos.getY());
        }else{
            return false;
        }
    }

    public Location addPosColModification(int colChange) {
        Pos newPos = pos.addYModification(colChange);
        return new Location(newPos, size);
    }

    public Location addPosRowModification(int rowChange) {
        Pos newPos = pos.addXModification(rowChange);
        return new Location(newPos, size);
    }

    public Pos getLeftBottomPos() {
        return new Pos(pos.getX() + size.getHeight(), pos.y);
    }
}
