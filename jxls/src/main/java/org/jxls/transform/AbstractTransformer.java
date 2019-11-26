package org.jxls.transform;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.jxls.common.AreaRef;
import org.jxls.common.CellData;
import org.jxls.common.CellRef;
import org.jxls.common.ImageType;
import org.jxls.common.RowData;
import org.jxls.common.SheetData;
import org.jxls.common.Size;

/**
 * Base transformer class providing basic implementation for some of the {@link Transformer} interface methods
 *
 * @author Leonid Vysochyn
 */
public abstract class AbstractTransformer implements Transformer {
    private boolean ignoreColumnProps = false;
    private boolean ignoreRowProps = false;
    protected final Map<String, SheetData> sheetMap = new LinkedHashMap<>();
    private TransformationConfig transformationConfig = new TransformationConfig();
    private boolean evaluateFormulas = false;

    @Override
    public List<CellRef> getTargetCellRef(CellRef cellRef) {
        CellData cellData = getCellData(cellRef);
        if (cellData != null) {
            return cellData.getTargetPos();
        } else {
            return new ArrayList<CellRef>();
        }
    }

    @Override
    public void resetTargetCellRefs() {
        for (SheetData sheetData : sheetMap.values()) {
            for (int i = 0; i < sheetData.getNumberOfRows(); i++) {
                RowData rowData = sheetData.getRowData(i);
                if (rowData != null) {
                    for (int j = 0; j < rowData.getNumberOfCells(); j++) {
                        CellData cellData = rowData.getCellData(j);
                        if (cellData != null) {
                            cellData.resetTargetPos();
                        }
                    }
                }
            }
        }
    }

    @Override
    public CellData getCellData(CellRef cellRef) {
        if (cellRef == null || cellRef.getSheetName() == null) {
            return null;
        }
        SheetData sheetData = sheetMap.get(cellRef.getSheetName());
        if (sheetData == null) {
            return null;
        }
        RowData rowData = sheetData.getRowData(cellRef.getRow());
        if (rowData == null) {
            return null;
        }
        return rowData.getCellData(cellRef.getCol());
    }

    public boolean isIgnoreColumnProps() {
        return ignoreColumnProps;
    }

    public void setIgnoreColumnProps(boolean ignoreColumnProps) {
        this.ignoreColumnProps = ignoreColumnProps;
    }

    public boolean isIgnoreRowProps() {
        return ignoreRowProps;
    }

    public void setIgnoreRowProps(boolean ignoreRowProps) {
        this.ignoreRowProps = ignoreRowProps;
    }

    @Override
    public TransformationConfig getTransformationConfig() {
        return transformationConfig;
    }

    @Override
    public void setTransformationConfig(TransformationConfig transformationConfig) {
        this.transformationConfig = transformationConfig;
    }

    @Override
    public Set<CellData> getFormulaCells() {
        Set<CellData> formulaCells = new HashSet<CellData>();
        for (SheetData sheetData : sheetMap.values()) {
            for (int i = 0; i < sheetData.getNumberOfRows(); i++) {
                RowData rowData = sheetData.getRowData(i);
                if (rowData != null) {
                    for (int j = 0; j < rowData.getNumberOfCells(); j++) {
                        CellData cellData = rowData.getCellData(j);
                        if (cellData != null && cellData.isFormulaCell()) {
                            formulaCells.add(cellData);
                        }
                    }
                }
            }
        }
        return formulaCells;
    }

    @Override
    public boolean deleteSheet(String sheetName) {
        if (sheetMap.containsKey(sheetName)) {
            sheetMap.remove(sheetName);
            return true;
        } else {
            return false;
        }
    }

    @Override
    public void adjustTableSize(CellRef ref, Size size) {
    }

    @Override
    public void dispose() {
    }

    @Override
    public void mergeCells(CellRef ref, int rows, int cols) {
        throw new UnsupportedOperationException("mergeCells operation is not implemented in the " + this.getClass().getName());
    }

    @Override
    public void writeButNotCloseStream() throws IOException {
        throw new UnsupportedOperationException("writeButNotCloseStream operation is not implemented in the " + this.getClass().getName());
    }

    @Override
    public void addImage(AreaRef areaRef, byte[] imageBytes, ImageType imageType, Double scaleX, Double scaleY) {
        throw new UnsupportedOperationException("addImage operation is not implemented in the " + this.getClass().getName());
    }

    @Override
    public boolean isEvaluateFormulas() {
        return evaluateFormulas;
    }

    /**
     * This option does not work for streaming.
     * 
     * @param evaluateFormulas false: formulas will be evaluated by MS Excel when opening the Excel file;
     * true (default): evaluate formulas before writing.
     */
    @Override
    public void setEvaluateFormulas(boolean evaluateFormulas) {
        this.evaluateFormulas = evaluateFormulas;
    }

    @Override
    public boolean isForwardOnly() {
        return false;
    }
}
