package org.jxls.common;

import org.jxls.util.CellRefUtil;

/**
 * Represents cell reference
 * 
 * @author Leonid Vysochyn
 */
public class CellRef implements Comparable<CellRef>{
    static final CellRef NONE = new CellRef("NONE", -1, -1);

    private int col;
    private int row;
    private String sheetName;
    private boolean isColAbs;
    private boolean isRowAbs;
    private boolean ignoreSheetNameInFormat = false;

    public CellRef(String sheetName, int row, int col) {
        this.sheetName = sheetName;
        this.row = row;
        this.col = col;
    }

    public CellRef(int row, int col) {
        this(null, row, col);
    }
    
    public CellRef(String cellRef) {
        if (cellRef.endsWith("#REF!")) {
            throw new IllegalArgumentException("Cell reference invalid: " + cellRef);
        }

        String[] parts = CellRefUtil.separateRefParts(cellRef);
        sheetName = parts[0];
        String colRef = parts[1];
        if (colRef.length() < 1) {
            throw new IllegalArgumentException("Invalid formula cell reference: '" + cellRef + "'");
        }
        isColAbs = colRef.charAt(0) == '$';
        if (isColAbs) {
            colRef = colRef.substring(1);
        }
        col = CellRefUtil.convertColStringToIndex(colRef);

        String rowRef = parts[2];
        if (rowRef.length() < 1) {
            throw new IllegalArgumentException("Invalid formula cell reference: '" + cellRef + "'");
        }
        isRowAbs = rowRef.charAt(0) == '$';
        if (isRowAbs) {
            rowRef = rowRef.substring(1);
        }
        row = Integer.parseInt(rowRef) - 1; // -1 to convert 1-based to zero-based
    }

    public String getCellName() {
        StringBuilder sb = new StringBuilder(32);
        if (sheetName != null && !ignoreSheetNameInFormat) {
            CellRefUtil.appendFormat(sb, sheetName);
            sb.append(CellRefUtil.SHEET_NAME_DELIMITER);
        }
        appendCellReference(sb);
        return sb.toString();
    }

    public String getFormattedSheetName() {
        StringBuilder sb = new StringBuilder(32);
        CellRefUtil.appendFormat(sb, sheetName);
        return sb.toString();
    }

    public String getSheetName() {
        return sheetName;
    }

    public CellRef setSheetName(String sheetName) {
        this.sheetName = sheetName;
        return this;
    }

    public boolean isIgnoreSheetNameInFormat() {
        return ignoreSheetNameInFormat;
    }

    public CellRef setIgnoreSheetNameInFormat(boolean ignoreSheetNameInFormat) {
        this.ignoreSheetNameInFormat = ignoreSheetNameInFormat;
        return this;
    }

    /**
     * Appends cell reference with '$' markers for absolute values as required.
     * Sheet name is not included.
     * @param sb
     */
    StringBuilder appendCellReference(StringBuilder sb) {
        if (isColAbs) {
            sb.append(CellRefUtil.ABSOLUTE_REFERENCE_MARKER);
        }
        sb.append(CellRefUtil.convertNumToColString(col));
        if (isRowAbs) {
            sb.append(CellRefUtil.ABSOLUTE_REFERENCE_MARKER);
        }
        sb.append(row + 1);
        return sb;
    }

    public int getCol() {
        return col;
    }

    public CellRef setCol(int col) {
        this.col = col;
        return this;
    }

    public int getRow() {
        return row;
    }

    public CellRef setRow(int row) {
        this.row = row;
        return this;
    }

    public boolean isColAbs() {
        return isColAbs;
    }

    public CellRef setIsColAbs(boolean isColAbs) {
        this.isColAbs = isColAbs;
        return this;
    }

    public boolean isRowAbs() {
        return isRowAbs;
    }

    public CellRef setIsRowAbs(boolean isRowAbs) {
        this.isRowAbs = isRowAbs;
        return this;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) {
            return true;
        }
        if (o == null || getClass() != o.getClass()) {
            return false;
        }
        CellRef cellRef = (CellRef) o;
        return !(col != cellRef.col
              || row != cellRef.row
              || (sheetName == null ? cellRef.sheetName != null : !sheetName.equals(cellRef.sheetName)));
    }

    @Override
    public int hashCode() {
        int result = col;
        result = 31 * result + row;
        result = 31 * result + (sheetName != null ? sheetName.hashCode() : 0);
        return result;
    }
    
    public String toString(boolean ignoreSheetName) {
        boolean currentIgnoreSheetValue = ignoreSheetNameInFormat;
        ignoreSheetNameInFormat = ignoreSheetName;
        String result = getCellName();
        ignoreSheetNameInFormat = currentIgnoreSheetValue;
        return result;
    }

    public boolean isValid() {
        return col >= 0 && row >= 0;
    }

    @Override
    public String toString() {
        return getCellName();
    }

    @Override
    public int compareTo(CellRef that) {
        if (this == that) {
            return 0;
        }
        if (col < that.col) {
            return -1;
        }
        if (col > that.col) {
            return 1;
        }
        if (row < that.row) {
            return -1;
        }
        if (row > that.row) {
            return 1;
        }
        return 0;
    }
}
