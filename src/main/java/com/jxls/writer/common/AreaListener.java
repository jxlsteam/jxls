package com.jxls.writer.common;

import com.jxls.writer.area.Area;

/**
 * @author Leonid Vysochyn
 *         Date: 2/16/12 4:45 PM
 */
public interface AreaListener {
    void beforeApplyAtCell(CellRef cellRef, Context context);
    void afterApplyAtCell(CellRef cellRef, Context context);
    void beforeTransformCell(CellRef srcCell, CellRef targetCell, Context context);
    void afterTransformCell(CellRef srcCell, CellRef targetCell, Context context);
}
