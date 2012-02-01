package com.jxls.writer.transform.poi;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * @author Leonid Vysochyn
 *         Date: 2/1/12 12:03 PM
 */
public class SheetData {
    
    String sheetName;
    int[] columnWidth;
    RowData[] rowData;
    
    public static SheetData createSheetData(Sheet sheet){
        SheetData sheetData = new SheetData();
        sheetData.sheetName = sheet.getSheetName();
        sheetData.columnWidth = new int[256];
        for(int i = 0; i < 256; i++){
            sheetData.columnWidth[i] = sheet.getColumnWidth(i);
        }
        int numberOfRows = sheet.getLastRowNum() + 1;
        sheetData.rowData = new RowData[numberOfRows];
        for(int i = 0; i < numberOfRows; i++){
            sheetData.rowData[i] = RowData.createRowData(sheet.getRow(i));
        }
        return sheetData;
    }

    public String getSheetName() {
        return sheetName;
    }

    public int getColumnWidth(int col) {
        return columnWidth[col];
    }
    
    public RowData getRowData(int rowIndex){
        return rowData[rowIndex];
    }

}
