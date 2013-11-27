package com.jxls.plus.transform;

import com.jxls.plus.common.*;

import java.io.IOException;
import java.util.List;
import java.util.Set;

/**
 * Defines interface methods for excel operations
 * @author Leonid Vysochyn
 *         Date: 1/23/12
 */
public interface Transformer {
    Context createInitialContext();
    void transform(CellRef srcCellRef, CellRef targetCellRef, Context context);
    void setFormula(CellRef cellRef, String formulaString);
    Set<CellData> getFormulaCells();
    CellData getCellData(CellRef cellRef);
    List<CellRef> getTargetCellRef(CellRef cellRef);
    void resetTargetCellRefs();
    void clearCell(CellRef cellRef);
    List<CellData> getCommentedCells();
    void addImage(AreaRef areaRef, byte[] imageBytes, ImageType imageType);
    void write() throws IOException;
}
