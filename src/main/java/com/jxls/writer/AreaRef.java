package com.jxls.writer;

/**
 * @author Leonid Vysochyn
 *         Date: 2/13/12 12:41 PM
 */
public class AreaRef {
    CellRef firstCellRef;
    CellRef lastCellRef;

    public AreaRef(CellRef firstCellRef, CellRef lastCellRef) {
        this.firstCellRef = firstCellRef;
        this.lastCellRef = lastCellRef;
    }
    
    public AreaRef(String areaRef){
        String[] parts = CellRefUtil.separateAreaRefs( areaRef );
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

}
