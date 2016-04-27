package org.jxls.transform;

import org.jxls.common.*;

import java.io.IOException;
import java.util.List;
import java.util.Set;

/**
 * Defines interface methods for excel operations
 *
 * @author Leonid Vysochyn
 *         Date: 1/23/12
 */
public interface Transformer {

    void transform(CellRef srcCellRef, CellRef targetCellRef, Context context, boolean updateRowHeight);

    void setFormula(CellRef cellRef, String formulaString);

    Set<CellData> getFormulaCells();

    CellData getCellData(CellRef cellRef);

    List<CellRef> getTargetCellRef(CellRef cellRef);

    void resetTargetCellRefs();

    void resetArea(AreaRef areaRef);

    void clearCell(CellRef cellRef);

    List<CellData> getCommentedCells();

    void addImage(AreaRef areaRef, byte[] imageBytes, ImageType imageType);

    void write() throws IOException;

    TransformationConfig getTransformationConfig();

    void setTransformationConfig(TransformationConfig transformationConfig);

    boolean deleteSheet(String sheetName);

    void setHidden(String sheetName, boolean hidden);

    void updateRowHeight(String srcSheetName, int srcRowNum, String targetSheetName, int targetRowNum);
}
