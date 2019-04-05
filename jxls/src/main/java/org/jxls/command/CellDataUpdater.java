package org.jxls.command;

import org.jxls.common.CellData;
import org.jxls.common.CellRef;
import org.jxls.common.Context;

/**
 * Interface for updating {@link CellData}
 * 
 * @see UpdateCellCommand
 */
public interface CellDataUpdater {

    void updateCellData(CellData cellData, CellRef targetCell, Context context);
}
