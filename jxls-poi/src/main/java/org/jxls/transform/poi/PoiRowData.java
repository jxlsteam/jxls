package org.jxls.transform.poi;

import org.apache.poi.ss.usermodel.Row;
import org.jxls.common.CellData;
import org.jxls.common.CellRef;
import org.jxls.common.RowData;

/**
 * Row data wrapper for POI row
 * 
 * @author Leonid Vysochyn
 */
public class PoiRowData extends RowData {
    private Row row;
    private PoiSheetData poiSheetData;

    public static RowData createRowData(PoiSheetData poiSheetData, Row row, PoiTransformer transformer) {
        if (row == null) {
            return null;
        }
        PoiRowData rowData = new PoiRowData();
        rowData.poiSheetData = poiSheetData;
        rowData.setTransformer(transformer);
        rowData.row = row;
        rowData.height = row.getHeight();
        int numberOfCells = row.getLastCellNum();
        for (int cellIndex = 0; cellIndex < numberOfCells; cellIndex++) {
            org.apache.poi.ss.usermodel.Cell cell = row.getCell(cellIndex);
            if (cell != null) {
                CellData cellData = PoiCellData.createCellData(rowData, new CellRef(row.getSheet().getSheetName(), row.getRowNum(), cellIndex), cell);
                cellData.setTransformer(transformer);
                rowData.addCellData(cellData);
            } else {
                rowData.addCellData(null);
            }
        }
        return rowData;
    }

    public Row getRow() {
        return row;
    }

    public PoiSheetData getPoiSheetData() {
        return poiSheetData;
    }
}
