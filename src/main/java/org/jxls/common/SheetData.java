package org.jxls.common;

import org.jxls.transform.Transformer;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * Represents an excel sheet data holder
 * @author Leonid Vysochyn
 */
public class SheetData implements Iterable<RowData> {
    protected String sheetName;
    protected int[] columnWidth;
    protected List<RowData> rowDataList = new ArrayList<RowData>();

    private Transformer transformer;

    public Transformer getTransformer() {
        return transformer;
    }

    public void setTransformer(Transformer transformer) {
        this.transformer = transformer;
    }

    public int getNumberOfRows(){
        return rowDataList.size();
    }

    public String getSheetName() {
        return sheetName;
    }

    public int getColumnWidth(int col) {
        return columnWidth[col];
    }

    public RowData getRowData(int row){
        if(row < rowDataList.size() ) return rowDataList.get(row);
        else return null;
    }

    public void addRowData(RowData rowData){
        rowDataList.add(rowData);
    }

    public Iterator<RowData> iterator() {
        return rowDataList.iterator();
    }
}
