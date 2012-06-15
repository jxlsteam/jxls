package com.jxls.writer.transform;

import com.jxls.writer.common.*;

import java.util.List;
import java.util.Set;

/**
 * @author Leonid Vysochyn
 *         Date: 1/23/12 1:24 PM
 */
public interface Transformer {
    void transform(CellRef srcCellRef, CellRef targetCellRef, Context context);
    void setFormula(CellRef cellRef, String formulaString);
    Set<CellData> getFormulaCells();
    CellData getCellData(CellRef cellRef);
    List<CellRef> getTargetCellRef(CellRef cellRef);
    void resetTargetCellRefs();
    void clearCell(CellRef cellRef);
    List<CellData> getCommentedCells();
    void addImage(AreaRef areaRef, int imageIdx);
    void addImage(AreaRef areaRef, byte[] imageBytes, ImageType imageType);
}
