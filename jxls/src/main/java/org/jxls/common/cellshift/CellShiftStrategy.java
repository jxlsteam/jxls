package org.jxls.common.cellshift;

import org.jxls.common.CellRef;

/**
 * Defines cell shift strategy
 */
public interface CellShiftStrategy {

    boolean requiresColShifting(CellRef cell, int startRow, int endRow, int startColShift);

    boolean requiresRowShifting(CellRef cell, int startCol, int endCol, int startRowShift);
}
