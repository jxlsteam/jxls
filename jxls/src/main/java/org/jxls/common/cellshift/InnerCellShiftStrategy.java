package org.jxls.common.cellshift;

import org.jxls.common.CellRef;

/**
 * Shifts only cells directly impacted by the cells being shifted
 * E.g. if a cell in column X is being shifted down it shifts all the other cells
 * located below the shifted cell in the column X
 */
public class InnerCellShiftStrategy implements CellShiftStrategy {

    @Override
    public boolean requiresColShifting(CellRef cell, int startRow, int endRow, int startColShift) {
        return cell != null && cell.getCol() > startColShift && cell.getRow() >= startRow && cell.getRow() <= endRow;
    }

    @Override
    public boolean requiresRowShifting(CellRef cell, int startCol, int endCol, int startRowShift) {
        return cell != null && cell.getRow() > startRowShift && cell.getCol() >= startCol && cell.getCol() <= endCol;
    }
}
