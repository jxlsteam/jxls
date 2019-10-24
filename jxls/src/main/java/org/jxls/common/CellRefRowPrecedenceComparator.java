package org.jxls.common;

import java.util.Comparator;

/**
 * Cell comparator used to order cell by rows first and then by columns
 * 
 * @author Leonid Vysochyn
 */
public class CellRefRowPrecedenceComparator implements Comparator<CellRef> {

    @Override
    public int compare(CellRef cellRef1, CellRef cellRef2) {
        if (cellRef1 == cellRef2) {
            return 0;
        }
        if (cellRef1 == null) {
            return 1;
        }
        if (cellRef2 == null) {
            return -1;
        }
        if (cellRef1.getSheetName() != null && cellRef2.getSheetName() != null) {
            int sheetNameCompared = cellRef1.getSheetName().compareTo(cellRef2.getSheetName());
            if (sheetNameCompared != 0) {
                return sheetNameCompared;
            }
        } else {
            if (cellRef1.getSheetName() != null || cellRef2.getSheetName() != null) {
                return cellRef1.getSheetName() != null ? -1 : 1;
            }
        }
        if (cellRef1.getRow() < cellRef2.getRow()) {
            return -1;
        }
        if (cellRef1.getRow() > cellRef2.getRow()) {
            return 1;
        }
        if (cellRef1.getCol() < cellRef2.getCol()) {
            return -1;
        }
        if (cellRef1.getCol() > cellRef2.getCol()) {
            return 1;
        }
        return 0;
    }
}
