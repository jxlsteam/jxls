package org.jxls.transform;

import org.jxls.common.*;
import org.jxls.expression.ExpressionEvaluator;
import org.jxls.expression.JexlExpressionEvaluator;

import java.util.*;
import java.util.regex.Pattern;

/**
 * Base transformer class providing basic implementation for some of the {@link Transformer} interface methods
 * @author Leonid Vysochyn
 *         Date: 2/6/12
 */
public abstract class AbstractTransformer implements Transformer {
    boolean ignoreColumnProps = false;
    boolean ignoreRowProps = false;
    protected Map<String, SheetData> sheetMap = new LinkedHashMap<String, SheetData>();

    TransformationConfig transformationConfig = new TransformationConfig();

    public AbstractTransformer() {
    }

    @Override
    public Context createInitialContext() {
        return new Context();
    }

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
}
