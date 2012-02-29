package com.jxls.writer.common;

import com.jxls.writer.util.CellRefUtil;

/**
 * @author Leonid Vysochyn
 *         Date: 2/13/12 12:41 PM
 */
public class AreaRef {
    CellRef firstCellRef;
    CellRef lastCellRef;

    public AreaRef(CellRef firstCellRef, CellRef lastCellRef) {
        if( (firstCellRef.getSheetName() == null && lastCellRef.getSheetName() == null) ||
                (firstCellRef.getSheetName() != null &&
                firstCellRef.getSheetName().equalsIgnoreCase(lastCellRef.getSheetName()) )){
            this.firstCellRef = firstCellRef;
            this.lastCellRef = lastCellRef;
        }else{
            throw new IllegalArgumentException("Cannot create area from specified cell references " + firstCellRef + ", " + lastCellRef);
        }
    }
    
    public AreaRef(CellRef cellRef, Size size){
        firstCellRef = cellRef;
        lastCellRef = new CellRef( cellRef.getSheetName(), cellRef.getRow() + size.getHeight() - 1, cellRef.getCol() + size.getWidth() - 1);
    }
    
    public AreaRef(String areaRef){
        String[] parts = CellRefUtil.separateAreaRefs(areaRef);
        String part0 = parts[0];
        if (parts.length == 1) {
            firstCellRef = new CellRef(part0);
            lastCellRef = firstCellRef;
            return;
        }
        if (parts.length != 2) {
            throw new IllegalArgumentException("Bad area ref '" + areaRef + "'");
        }

        String part1 = parts[1];
        if (CellRefUtil.isPlainColumn(part0) || CellRefUtil.isPlainColumn(part1)) {
            throw new IllegalArgumentException("Plain column references are not currently supported");
        } else {
            firstCellRef = new CellRef(part0);
            lastCellRef = new CellRef(part1);
        }
    }
    
    public String getSheetName(){
        return firstCellRef.getSheetName();
    }

    public CellRef getFirstCellRef() {
        return firstCellRef;
    }

    public CellRef getLastCellRef() {
        return lastCellRef;
    }
    
    public Size getSize(){
        if( firstCellRef == null || lastCellRef == null ) return Size.ZERO_SIZE;
        return new Size(lastCellRef.getCol() - firstCellRef.getCol() + 1, lastCellRef.getRow() - firstCellRef.getRow() + 1);
    }
    
    boolean contains(CellRef cellRef){
        if( (getSheetName() == null && cellRef.getSheetName() == null) ||
                getSheetName() != null && getSheetName().equalsIgnoreCase(cellRef.getSheetName())){
            return (cellRef.getRow() >= firstCellRef.getRow() && cellRef.getCol() >= firstCellRef.getCol() &&
                    cellRef.getRow() <= lastCellRef.getRow() && cellRef.getCol() <= lastCellRef.getCol());
        }else{
            return false;
        }
    }
    
    public boolean contains(AreaRef areaRef){
        if( areaRef == null ) return true;
        if( (getSheetName() == null && areaRef.getSheetName() == null) ||
                getSheetName().equalsIgnoreCase(areaRef.getSheetName())){
            return contains(areaRef.getFirstCellRef()) && contains( areaRef.getLastCellRef() );
        }else{
            return false;
        }
    }

    @Override
    public String toString() {
        return firstCellRef.toString() + ":" + lastCellRef.toString(true);
    }
}
