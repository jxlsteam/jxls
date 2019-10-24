package org.jxls.common;

import org.jxls.util.CellRefUtil;

/**
 * Defines area bounds
 * 
 * @author Leonid Vysochyn
 */
public class AreaRef {
    private CellRef firstCellRef;
    private CellRef lastCellRef;
    private int startRow;
    private int startCol;
    private int endRow;
    private int endCol;

    public AreaRef(CellRef firstCellRef, CellRef lastCellRef) {
        // TODO The following if-condition is used 3 times in this class. Write testcase and refactor!
        if ((firstCellRef.getSheetName() == null && lastCellRef.getSheetName() == null)
                || (firstCellRef.getSheetName() != null && firstCellRef.getSheetName().equalsIgnoreCase(lastCellRef.getSheetName()))) {
            this.firstCellRef = firstCellRef;
            this.lastCellRef = lastCellRef;
            updateStartEndRowCol();
        } else {
            throw new IllegalArgumentException("Cannot create area from specified cell references " + firstCellRef + ", " + lastCellRef);
        }
    }
    
    public AreaRef(CellRef cellRef, Size size) {
        firstCellRef = cellRef;
        lastCellRef = new CellRef(cellRef.getSheetName(),
                cellRef.getRow() + size.getHeight() - 1,
                cellRef.getCol() + size.getWidth() - 1);
        updateStartEndRowCol();
    }
    
    public AreaRef(String areaRef) {
        String[] parts = CellRefUtil.separateAreaRefs(areaRef);
        String part0 = parts[0];
        if (parts.length == 1) {
            firstCellRef = new CellRef(part0);
            lastCellRef = firstCellRef;
            updateStartEndRowCol();
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
        updateStartEndRowCol();
    }
    
    private void updateStartEndRowCol() {
        startRow = firstCellRef.getRow();
        startCol = firstCellRef.getCol();
        endRow = lastCellRef.getRow();
        endCol = lastCellRef.getCol();
    }

    public String getSheetName() {
        return firstCellRef.getSheetName();
    }

    public CellRef getFirstCellRef() {
        return firstCellRef;
    }

    public CellRef getLastCellRef() {
        return lastCellRef;
    }

    public Size getSize() {
        if (firstCellRef == null || lastCellRef == null) {
            return Size.ZERO_SIZE;
        }
        return new Size(endCol - startCol + 1, endRow - startRow + 1);
    }
    
    public boolean contains(CellRef cellRef) {
        String sheetName = getSheetName();
        String otherSheetName = cellRef.getSheetName();
        if ((sheetName == null && otherSheetName == null)
                || (sheetName != null && sheetName.equalsIgnoreCase(otherSheetName))) {

            return     cellRef.getRow() >= startRow
                    && cellRef.getCol() >= startCol
                    && cellRef.getRow() <= endRow
                    && cellRef.getCol() <= endCol;
        } else {
            return false;
        }
    }

    public boolean contains(int row, int col) {
        return     row >= startRow && row <= endRow
                && col >= startCol && col <= endCol;
    }
    
    public boolean contains(AreaRef areaRef) {
        if (areaRef == null) {
            return true;
        }
        if ((getSheetName() == null && areaRef.getSheetName() == null)
                || (getSheetName() != null && getSheetName().equalsIgnoreCase(areaRef.getSheetName()))) {
            return contains(areaRef.getFirstCellRef()) && contains(areaRef.getLastCellRef());
        } else {
            return false;
        }
    }

    @Override
    public String toString() {
        return firstCellRef.toString() + ":" + lastCellRef.toString(true);
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) {
            return true;
        }
        if (o == null || getClass() != o.getClass()) {
            return false;
        }

        AreaRef areaRef = (AreaRef) o;
        if (firstCellRef != null ? !firstCellRef.equals(areaRef.firstCellRef) : areaRef.firstCellRef != null) {
            return false;
        }
        if (lastCellRef != null ? !lastCellRef.equals(areaRef.lastCellRef) : areaRef.lastCellRef != null) {
            return false;
        }
        return true;
    }

    @Override
    public int hashCode() {
        int result = firstCellRef != null ? firstCellRef.hashCode() : 0;
        result = 31 * result + (lastCellRef != null ? lastCellRef.hashCode() : 0);
        return result;
    }
}
