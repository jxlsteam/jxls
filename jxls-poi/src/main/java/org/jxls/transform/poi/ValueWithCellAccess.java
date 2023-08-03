package org.jxls.transform.poi;

import org.apache.poi.ss.usermodel.Cell;

public interface ValueWithCellAccess {

    Object getValue();

    void applyBefore(Cell cell);
    
    void applyAfter(Cell cell);
}
