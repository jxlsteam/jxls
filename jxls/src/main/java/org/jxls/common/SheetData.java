package org.jxls.common;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.jxls.transform.Transformer;

/**
 * Represents an Excel sheet data holder
 * 
 * @author Leonid Vysochyn
 */
public class SheetData implements Iterable<RowData> {
    protected String sheetName;
    protected int[] columnWidth;
    protected final List<RowData> rowDataList = new ArrayList<RowData>();
    private Transformer transformer;

    public Transformer getTransformer() {
        return transformer;
    }

    public void setTransformer(Transformer transformer) {
        this.transformer = transformer;
    }

    public int getNumberOfRows() {
        return rowDataList.size();
    }

    public String getSheetName() {
        return sheetName;
    }

    public int getColumnWidth(int col) {
        return columnWidth[col];
    }

    public RowData getRowData(int row) {
        if (row < rowDataList.size()) {
            return rowDataList.get(row);
        } else {
            return null;
        }
    }

    public void addRowData(RowData rowData) {
        rowDataList.add(rowData);
    }

    @Override
    public Iterator<RowData> iterator() {
        return rowDataList.iterator();
    }
}
