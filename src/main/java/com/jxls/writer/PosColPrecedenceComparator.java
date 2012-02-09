package com.jxls.writer;

import java.util.Comparator;

/**
 * @author Leonid Vysochyn
 *         Date: 2/8/12 5:22 PM
 */
public class PosColPrecedenceComparator implements Comparator<Pos>{

    public int compare(Pos pos1, Pos pos2) {
        if( pos1 == pos2 ) return 0;
        if( pos1 == null) return 1;
        if( pos2 == null) return -1;
        if( pos1.getSheetName() != null && pos2.getSheetName() != null ){
            int sheetNameCompared = pos1.getSheetName().compareTo(pos2.getSheetName());
            if( sheetNameCompared != 0 ) return sheetNameCompared;
        }
        if(pos1.getSheetName() != null || pos2.getSheetName() != null){
            if( pos1.getSheetName() != null ) return -1;
            else return 1;
        }
        if( pos1.getCol() < pos2.getCol() ) return -1;
        if( pos1.getCol() > pos2.getCol() ) return 1;
        if( pos1.getRow() < pos2.getRow() ) return -1;
        if( pos1.getRow() > pos2.getRow() ) return 1;
        return 0;
    }
}
