package org.jxls.transform;

import java.io.IOException;
import java.io.OutputStream;
import java.util.List;
import java.util.Set;

import org.jxls.common.AreaRef;
import org.jxls.common.CellData;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.Size;
import org.jxls.logging.JxlsLogger;

/**
 * Defines interface methods for Excel operations
 *
 * @author Leonid Vysochyn
 */
public interface Transformer {

    void transform(CellRef srcCellRef, CellRef targetCellRef, Context context, boolean updateRowHeight);

    void setOutputStream(OutputStream outputStream);
    
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

    boolean deleteSheet(String sheetName);

    void setHidden(String sheetName, boolean hidden);

    void updateRowHeight(String srcSheetName, int srcRowNum, String targetSheetName, int targetRowNum);

    void adjustTableSize(CellRef ref, Size size);

    /**
     * This property is used to recalculate all formulas before saving the workbook.
     * This property is set to true if you don't open the file with MS Excel and just read it (e.g. with a unit test).
     * The following documentation is POI specific.
     * 
     * @return true: calls <code>workbook.getCreationHelper().createFormulaEvaluator().evaluateAll()</code>
     * before writing the workbook. Please have a look at the POI documentation for more details.
     * This does not work for streaming.
     * Please be aware that POI supports only a subset of Excel formulas.
     * If an unsupported formula is in the template the evaluation will fail.
     * <p>false (default): do nothing (hopefully MS Excel will recalculate all formulas while opening the file)</p>
     */
    boolean isEvaluateFormulas();

    /**
     * This property is used to recalculate all formulas before saving the workbook.
     * This property is set to true if you don't open the file with MS Excel and just read it (e.g. with a unit test).
     * The following documentation is POI specific.
     * 
     * @param evaluateFormulas
     * true: calls <code>workbook.getCreationHelper().createFormulaEvaluator().evaluateAll()</code>
     * before writing the workbook. Please have a look at the POI documentation for more details.
     * This does not work for streaming.
     * Please be aware that POI supports only a subset of Excel formulas.
     * If an unsupported formula is in the template the evaluation will fail.
     * <p>false (default): do nothing (hopefully MS Excel will recalculate all formulas while opening the file)</p>
     */
    void setEvaluateFormulas(boolean evaluateFormulas);

    /**
     * @return false: do nothing, true: activate recalculation when opening
     */
    boolean isFullFormulaRecalculationOnOpening();

     /**
      * If you set this option to true, all formulas will be recalculated when the file is opened in MS Excel.
      * This changes the Excel file. This only works once. 
      * @param fullFormulaRecalculationOnOpening false: do nothing (default value),
      *        true: activate recalculation when opening
      */
    void setFullFormulaRecalculationOnOpening(boolean fullFormulaRecalculationOnOpening);

    /**
     * @return true if the transformer can process cells only in a single pass
     */
    boolean isForwardOnly();

    /**
     * @return never null
     */
    JxlsLogger getLogger();
    
    /**
     * @param logger not null
     */
    void setLogger(JxlsLogger logger);
    
    void setIgnoreColumnProps(boolean ignoreColumnProps);
    
    void setIgnoreRowProps(boolean ignoreRowProps);
}
