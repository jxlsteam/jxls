package org.jxls.command;

import org.jxls.common.CellData;
import org.jxls.common.CellRef;
import org.jxls.common.Context;

/**
 * Created by leona_000 on 2017-03-16.
 */
public interface CellDataUpdater {
    void updateCellData(CellData cellData, CellRef targetCell, Context context);
}
