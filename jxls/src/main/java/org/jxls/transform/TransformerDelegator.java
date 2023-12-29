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
 * Decorator pattern for Transformer, in particular write() can be extended overriding beforeWrite()
 */
public class TransformerDelegator implements Transformer {
    protected final Transformer transformer;
    
    public TransformerDelegator(Transformer transformer) {
        if (transformer == null) {
            throw new IllegalArgumentException("transformer must not be null");
        }
        this.transformer = transformer;
    }

    @Override
    public void transform(CellRef srcCellRef, CellRef targetCellRef, Context context, boolean updateRowHeight) {
        transformer.transform(srcCellRef, targetCellRef, context, updateRowHeight);
    }

    @Override
    public void setOutputStream(OutputStream outputStream) {
        transformer.setOutputStream(outputStream);
    }

    @Override
	public void writeButNotCloseStream() throws IOException {
	    transformer.writeButNotCloseStream();
	}

	protected void beforeWrite() {
    }
    
    @Override
    public void write() throws IOException {
        beforeWrite();
        transformer.write();
    }

    @Override
    public void dispose() {
        transformer.dispose();
    }

    @Override
    public void setFormula(CellRef cellRef, String formulaString) {
        transformer.setFormula(cellRef, formulaString);
    }

    @Override
    public Set<CellData> getFormulaCells() {
        return transformer.getFormulaCells();
    }

    @Override
    public CellData getCellData(CellRef cellRef) {
        return transformer.getCellData(cellRef);
    }

    @Override
    public List<CellRef> getTargetCellRef(CellRef cellRef) {
        return transformer.getTargetCellRef(cellRef);
    }

    @Override
    public void resetTargetCellRefs() {
        transformer.resetTargetCellRefs();
    }

    @Override
    public void resetArea(AreaRef areaRef) {
        transformer.resetArea(areaRef);
    }

    @Override
    public void clearCell(CellRef cellRef) {
        transformer.clearCell(cellRef);
    }

    @Override
    public List<CellData> getCommentedCells() {
        return transformer.getCommentedCells();
    }

    @Override
    public boolean deleteSheet(String sheetName) {
        return transformer.deleteSheet(sheetName);
    }

    @Override
    public void setHidden(String sheetName, boolean hidden) {
        transformer.setHidden(sheetName, hidden);
    }

    @Override
    public void updateRowHeight(String srcSheetName, int srcRowNum, String targetSheetName, int targetRowNum) {
        transformer.updateRowHeight(srcSheetName, srcRowNum, targetSheetName, targetRowNum);
    }

    @Override
    public void adjustTableSize(CellRef ref, Size size) {
        transformer.adjustTableSize(ref, size);
    }

    @Override
    public boolean isEvaluateFormulas() {
        return transformer.isEvaluateFormulas();
    }

    @Override
    public void setEvaluateFormulas(boolean evaluateFormulas) {
        transformer.setEvaluateFormulas(evaluateFormulas);
    }

    @Override
    public boolean isFullFormulaRecalculationOnOpening() {
        return transformer.isFullFormulaRecalculationOnOpening();
    }

    @Override
    public void setFullFormulaRecalculationOnOpening(boolean fullFormulaRecalculationOnOpening) {
        transformer.setFullFormulaRecalculationOnOpening(fullFormulaRecalculationOnOpening);
    }

    @Override
    public boolean isForwardOnly() {
        return transformer.isForwardOnly();
    }

    @Override
	public void setLogger(JxlsLogger logger) {
        transformer.setLogger(logger);
    }

    @Override
    public JxlsLogger getLogger() {
        return transformer.getLogger();
    }

    @Override
	public void setIgnoreColumnProps(boolean ignoreColumnProps) {
		transformer.setIgnoreColumnProps(ignoreColumnProps);
	}

	@Override
	public void setIgnoreRowProps(boolean ignoreRowProps) {
		transformer.setIgnoreRowProps(ignoreRowProps);
	}
}
