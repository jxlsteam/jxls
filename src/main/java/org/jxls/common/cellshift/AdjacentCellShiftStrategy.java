package org.jxls.common.cellshift;

import org.jxls.common.CellRef;

/**
 * Created by Leonid Vysochyn on 07-Aug-15.
 */
public class AdjacentCellShiftStrategy implements CellShiftStrategy {
    @Override
    public boolean requiresColShifting(CellRef cell, int startRow, int endRow, int startColShift) {
        return cell != null && cell.getCol() > startColShift;
    }

    @Override
    public boolean requiresRowShifting(CellRef cell, int startCol, int endCol, int startRowShift) {
        return cell != null && cell.getRow() > startRowShift;
    }
}
