package org.jxls.transform;

import java.io.IOException;
import java.util.List;
import java.util.Set;

import org.jxls.common.AreaRef;
import org.jxls.common.CellData;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.ImageType;
import org.jxls.common.Size;

/**
 * Defines interface methods for Excel operations
 *
 * @author Leonid Vysochyn
 */
public interface Transformer {

    void setTransformationConfig(TransformationConfig transformationConfig);

    TransformationConfig getTransformationConfig();

    void transform(CellRef srcCellRef, CellRef targetCellRef, Context context, boolean updateRowHeight);

    /***
     * Writes Excel workbook to output stream but not close the stream
     * designed to use with ZipOutputStream or other OutputStream
     * for creates several xls files one time.
     * 
     * @throws IOException -
     */
    void writeButNotCloseStream() throws IOException;

    /**
     * Writes Excel workbook to output stream and disposes the workbook.
     * 
     * @throws IOException -
     */
    void write() throws IOException;
    
    /**
     * Must be called after use. write() calls this method.
     */
    void dispose();

    void setFormula(CellRef cellRef, String formulaString);

    Set<CellData> getFormulaCells();

    CellData getCellData(CellRef cellRef);

    /**
     * @param cellRef a source cell reference
     * @return a list of cell references into which the source cell was transformed
     */
    List<CellRef> getTargetCellRef(CellRef cellRef);

    void resetTargetCellRefs();

    void resetArea(AreaRef areaRef);

    void clearCell(CellRef cellRef);

    List<CellData> getCommentedCells();

    void addImage(AreaRef areaRef, byte[] imageBytes, ImageType imageType);

    void addImage(AreaRef areaRef, byte[] imageBytes, ImageType imageType, Double scaleX, Double scaleY);

    boolean deleteSheet(String sheetName);

    void setHidden(String sheetName, boolean hidden);

    void updateRowHeight(String srcSheetName, int srcRowNum, String targetSheetName, int targetRowNum);

    void adjustTableSize(CellRef ref, Size size);

    void mergeCells(CellRef ref, int rows, int cols);

    /**
     * @return false: formulas will be evaluated by MS Excel when opening the Excel file, true: evaluate formulas before writing
     */
    boolean isEvaluateFormulas();

    /**
     * @param evaluateFormulas false: formulas will be evaluated by MS Excel when opening the Excel file, true: evaluate formulas before writing
     */
    void setEvaluateFormulas(boolean evaluateFormulas);

    /**
     * @return true if the transformer can process cells only in a single pass
     */
    boolean isForwardOnly();
}
