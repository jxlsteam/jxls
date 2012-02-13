package com.jxls.writer;

import java.util.Comparator;

/**
 * @author Leonid Vysochyn
 *         Date: 2/8/12 5:22 PM
 */
public class PosRowPrecedenceComparator implements Comparator<CellRef>{

    public int compare(CellRef cellRef1, CellRef cellRef2) {
        if( cellRef1 == cellRef2) return 0;
        if( cellRef1 == null) return 1;
        if( cellRef2 == null) return -1;
        if( cellRef1.getSheetName() != null && cellRef2.getSheetName() != null ){
            int sheetNameCompared = cellRef1.getSheetName().compareTo(cellRef2.getSheetName());
            if( sheetNameCompared != 0 ) return sheetNameCompared;
        }else{
            if(cellRef1.getSheetName() != null || cellRef2.getSheetName() != null){
                if( cellRef1.getSheetName() != null ) return -1;
                else return 1;
            }
        }
        if( cellRef1.getRow() < cellRef2.getRow() ) return -1;
        if( cellRef1.getRow() > cellRef2.getRow() ) return 1;
        if( cellRef1.getCol() < cellRef2.getCol() ) return -1;
        if( cellRef1.getCol() > cellRef2.getCol() ) return 1;
        return 0;
    }
}
