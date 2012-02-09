package com.jxls.writer;

/**
 * @author Leonid Vysochyn
 *         Date: 1/25/12 1:42 PM
 */
public class Pos{

    int col;
    int row;
    int sheet;

    String sheetName;
    boolean isColAbs;
    boolean isRowAbs;

    public Pos(String sheetName, int row, int col) {
        this.sheetName = sheetName;
        this.row = row;
        this.col = col;
    }

    public Pos(int sheet, int row, int col) {
        this.sheet = sheet;
        this.row = row;
        this.col = col;
    }

    public Pos(int row, int col) {
        this(0, row, col);
    }
    
    public Pos(String cellRef){
        if(cellRef.endsWith("#REF!")) {
            throw new IllegalArgumentException("Cell reference invalid: " + cellRef);
        }

        String[] parts = CellRefUtil.separateRefParts(cellRef);
        sheetName = parts[0];
        String colRef = parts[1];
        if (colRef.length() < 1) {
            throw new IllegalArgumentException("Invalid Formula cell reference: '"+cellRef+"'");
        }
        isColAbs = colRef.charAt(0) == '$';
        if (isColAbs) {
            colRef=colRef.substring(1);
        }
        col = CellRefUtil.convertColStringToIndex(colRef);

        String rowRef=parts[2];
        if (rowRef.length() < 1) {
            throw new IllegalArgumentException("Invalid Formula cell reference: '"+cellRef+"'");
        }
        isRowAbs = rowRef.charAt(0) == '$';
        if (isRowAbs) {
            rowRef=rowRef.substring(1);
        }
        row = Integer.parseInt(rowRef)-1; // -1 to convert 1-based to zero-based
    }

    public String getCellName(){
        StringBuffer sb = new StringBuffer(32);
        if(sheetName != null) {
            CellRefUtil.appendFormat(sb, sheetName);
            sb.append(CellRefUtil.SHEET_NAME_DELIMITER);
        }
        appendCellReference(sb);
        return sb.toString();
    }

    public String getSheetName() {
        return sheetName;
    }

    /**
     * Appends cell reference with '$' markers for absolute values as required.
     * Sheet name is not included.
     */
    void appendCellReference(StringBuffer sb) {
        if(isColAbs) {
            sb.append(CellRefUtil.ABSOLUTE_REFERENCE_MARKER);
        }
        sb.append( CellRefUtil.convertNumToColString(col));
        if(isRowAbs) {
            sb.append(CellRefUtil.ABSOLUTE_REFERENCE_MARKER);
        }
        sb.append(row+1);
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

    @Override
    public String toString() {
        return getCellName();
    }

    public int compareTo(Pos pos) {
        if( this == pos ) return 0;
        if( pos == null ) return -1;
        if( pos.getSheet() < sheet ) return 1;
        if( pos.getSheet() > sheet ) return -1;
        if( pos.getCol() > col ) return -1;
        if( pos.getCol() < col ) return 1;
        if( pos.getRow() < row ) return 1;
        if( pos.getRow() > row ) return -1;
        return 0;
    }
}
