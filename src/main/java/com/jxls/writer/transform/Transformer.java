package com.jxls.writer.transform;

import com.jxls.writer.CellRef;
import com.jxls.writer.CellData;
import com.jxls.writer.command.Context;

import java.util.List;
import java.util.Set;

/**
 * @author Leonid Vysochyn
 *         Date: 1/23/12 1:24 PM
 */
public interface Transformer {
    void transform(CellRef cellRef, CellRef newCellRef, Context context);
    void setFormula(CellRef cellRef, String formulaString);
    Set<CellData> getFormulaCells();
    CellData getCellData(CellRef cellRef);
    List<CellRef> getTargetPos(CellRef cellRef);
    void resetTargetCells();
    int getSheetIndex(String sheetName);
}
