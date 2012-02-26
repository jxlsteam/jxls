package com.jxls.writer.transform.poi;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * @author Leonid Vysochyn
 *         Date: 2/1/12 12:03 PM
 */
public class SheetData implements Iterable<RowData>{
    String sheetName;
    int[] columnWidth;
    List<RowData> rowDataList = new ArrayList<RowData>();
    List<CellRangeAddress> mergedRegions = new ArrayList<CellRangeAddress>();
    
    public static SheetData createSheetData(Sheet sheet){
        SheetData sheetData = new SheetData();
        sheetData.sheetName = sheet.getSheetName();
        sheetData.columnWidth = new int[256];
        for(int i = 0; i < 256; i++){
            sheetData.columnWidth[i] = sheet.getColumnWidth(i);
        }
        int numberOfRows = sheet.getLastRowNum() + 1;
        for(int i = 0; i < numberOfRows; i++){
            sheetData.rowDataList.add(RowData.createRowData(sheet.getRow(i)));
        }
        for(int i = 0; i < sheet.getNumMergedRegions(); i++){
            CellRangeAddress region = sheet.getMergedRegion(i);
            sheetData.mergedRegions.add(region);
        }
        return sheetData;
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

    public List<CellRangeAddress> getMergedRegions() {
        return mergedRegions;
    }

    public Iterator<RowData> iterator() {
        return rowDataList.iterator();
    }
}
