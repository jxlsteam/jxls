package org.jxls.common;

/**
 * Defines callback methods to be called during area processing
 * @author Leonid Vysochyn
 *         Date: 2/16/12
 */
public interface AreaListener {
    void beforeApplyAtCell(CellRef cellRef, Context context);
    void afterApplyAtCell(CellRef cellRef, Context context);
    void beforeTransformCell(CellRef srcCell, CellRef targetCell, Context context);
    void afterTransformCell(CellRef srcCell, CellRef targetCell, Context context);
}
