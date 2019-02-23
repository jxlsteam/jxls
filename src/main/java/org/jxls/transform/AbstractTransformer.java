package org.jxls.transform;

import org.jxls.common.CellData;
import org.jxls.common.CellRef;
import org.jxls.common.RowData;
import org.jxls.common.SheetData;
import org.jxls.common.Size;

import java.util.ArrayList;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * Base transformer class providing basic implementation for some of the {@link Transformer} interface methods
 *
 * @author Leonid Vysochyn
 *         Date: 2/6/12
 */
public abstract class AbstractTransformer implements Transformer {
    private boolean ignoreColumnProps = false;
    private boolean ignoreRowProps = false;
    protected Map<String, SheetData> sheetMap = new LinkedHashMap<>();

    private TransformationConfig transformationConfig = new TransformationConfig();

    public List<CellRef> getTargetCellRef(CellRef cellRef) {
        CellData cellData = getCellData(cellRef);
        if (cellData != null) {
            return cellData.getTargetPos();
        } else {
            return new ArrayList<CellRef>();
        }
    }

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

    public CellData getCellData(CellRef cellRef) {
        if (cellRef == null || cellRef.getSheetName() == null) return null;
        SheetData sheetData = sheetMap.get(cellRef.getSheetName());
        if (sheetData == null) return null;
        RowData rowData = sheetData.getRowData(cellRef.getRow());
        if (rowData == null) return null;
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
}
