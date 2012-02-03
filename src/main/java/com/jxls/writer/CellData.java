package com.jxls.writer;

import com.jxls.writer.transform.poi.PoiCellData;

/**
 * @author Leonid Vysochyn
 *         Date: 2/3/12 12:18 PM
 */
public class CellData {
    protected String formula;
    protected int colIndex;
    protected int rowIndex;
    protected int sheetIndex;
    protected Object evaluationResult;
    protected Object cellOriginalValue;

    public CellData() {
    }

    public CellData(int sheetIndex, int rowIndex, int colIndex) {
        this.sheetIndex = sheetIndex;
        this.rowIndex = rowIndex;
        this.colIndex = colIndex;
    }

    public CellData(int rowIndex, int colIndex) {
        this.rowIndex = rowIndex;
        this.colIndex = colIndex;
    }

    public Object getCellValue(){
        return cellOriginalValue;
    }

    public int getSheetIndex() {
        return sheetIndex;
    }

    public void setSheetIndex(int sheetIndex) {
        this.sheetIndex = sheetIndex;
    }

    public int getRowIndex() {
        return rowIndex;
    }

    public void setRowIndex(int rowIndex) {
        this.rowIndex = rowIndex;
    }

    public int getColIndex() {
        return colIndex;
    }

    public void setColIndex(int colIndex) {
        this.colIndex = colIndex;
    }

    public String getFormula() {
        return formula;
    }

    public void setFormula(String formula) {
        this.formula = formula;
    }


    public boolean isUserFormula(String str) {
        return str.startsWith(PoiCellData.USER_FORMULA_PREFIX) && str.endsWith(PoiCellData.USER_FORMULA_SUFFIX);
    }
}
