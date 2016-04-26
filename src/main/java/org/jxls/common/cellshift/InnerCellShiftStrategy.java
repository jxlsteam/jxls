package org.jxls.common.cellshift;

import org.jxls.common.CellRef;

public class InnerCellShiftStrategy implements CellShiftStrategy{
    public boolean requiresColShifting(CellRef cell, int startRow, int endRow, int startColShift) {
        return cell != null && cell.getCol() > startColShift && cell.getRow() >= startRow && cell.getRow() <= endRow;
    }

    public boolean requiresRowShifting(CellRef cell, int startCol, int endCol, int startRowShift) {
        return cell != null && cell.getRow() > startRowShift && cell.getCol() >= startCol && cell.getCol() <= endCol;
    }
}