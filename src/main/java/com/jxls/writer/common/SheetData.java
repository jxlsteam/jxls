package com.jxls.writer.common;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * @author Leonid Vysochyn
 */
public class SheetData implements Iterable<RowData> {
    protected String sheetName;
    protected int[] columnWidth;
    protected List<RowData> rowDataList = new ArrayList<RowData>();

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
