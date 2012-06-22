package com.jxls.plus.transform.poi;

import com.jxls.plus.common.SheetData;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.List;

/**
 * Sheet data wrapper for POI sheet
 * @author Leonid Vysochyn
 *         Date: 2/1/12
 */
public class PoiSheetData extends SheetData {
    List<CellRangeAddress> mergedRegions = new ArrayList<CellRangeAddress>();
    Sheet sheet;
    
    public static PoiSheetData createSheetData(Sheet sheet){
        PoiSheetData sheetData = new PoiSheetData();
        sheetData.sheet = sheet;
        sheetData.sheetName = sheet.getSheetName();
        sheetData.columnWidth = new int[256];
        for(int i = 0; i < 256; i++){
            sheetData.columnWidth[i] = sheet.getColumnWidth(i);
        }
        int numberOfRows = sheet.getLastRowNum() + 1;
        for(int i = 0; i < numberOfRows; i++){
            sheetData.rowDataList.add(PoiRowData.createRowData(sheet.getRow(i)));
        }
        for(int i = 0; i < sheet.getNumMergedRegions(); i++){
            CellRangeAddress region = sheet.getMergedRegion(i);
            sheetData.mergedRegions.add(region);
        }
        return sheetData;
    }

    public List<CellRangeAddress> getMergedRegions() {
        return mergedRegions;
    }

    public Sheet getSheet() {
        return sheet;
    }
}
