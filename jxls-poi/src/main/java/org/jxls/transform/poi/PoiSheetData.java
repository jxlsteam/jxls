package org.jxls.transform.poi;

import org.jxls.common.RowData;
import org.jxls.common.SheetData;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.List;

/**
 * Sheet data wrapper for POI sheet
 * 
 * @author Leonid Vysochyn
 */
public class PoiSheetData extends SheetData {
    private List<CellRangeAddress> mergedRegions = new ArrayList<>();
    private Sheet sheet;

    public static PoiSheetData createSheetData(Sheet sheet, PoiTransformer transformer) {
        PoiSheetData sheetData = new PoiSheetData();
        sheetData.setTransformer(transformer);
        sheetData.sheet = sheet;
        sheetData.sheetName = sheet.getSheetName();
        int numberOfRows = sheet.getLastRowNum() + 1;
        int numberOfColumns = -1;
        for (int i = 0; i < numberOfRows; i++) {
            RowData rowData = PoiRowData.createRowData(sheet.getRow(i), transformer);
            sheetData.rowDataList.add(rowData);
            if (rowData != null && rowData.getNumberOfCells() > numberOfColumns) {
                numberOfColumns = rowData.getNumberOfCells();
            }
        }
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress region = sheet.getMergedRegion(i);
            sheetData.mergedRegions.add(region);
        }
        if (numberOfColumns > 0) {
            sheetData.columnWidth = new int[numberOfColumns];
            for (int i = 0; i < numberOfColumns; i++) {
                sheetData.columnWidth[i] = sheet.getColumnWidth(i);
            }
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
