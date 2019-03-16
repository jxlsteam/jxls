package org.jxls.common.cellshift;

import org.jxls.common.CellRef;

/**
 * Shifts not only cells impacted by shifting area but also the adjacent area cells
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
