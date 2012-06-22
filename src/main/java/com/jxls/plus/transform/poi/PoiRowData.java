package com.jxls.plus.transform.poi;

import com.jxls.plus.common.CellData;
import com.jxls.plus.common.CellRef;
import com.jxls.plus.common.RowData;
import org.apache.poi.ss.usermodel.Row;

/**
 * Row data wrapper for POI row
 * @author Leonid Vysochyn
 *         Date: 2/1/12
 */
public class PoiRowData extends RowData {
    Row row;

    public static RowData createRowData(Row row){
        if( row == null ) return null;
        PoiRowData rowData = new PoiRowData();
        rowData.row = row;
        rowData.height = row.getHeight();
        int numberOfCells = row.getLastCellNum();
        for(int cellIndex = 0; cellIndex < numberOfCells; cellIndex++){
            org.apache.poi.ss.usermodel.Cell cell = row.getCell(cellIndex);
            if(cell != null ){
                CellData cellData = PoiCellData.createCellData(new CellRef(row.getSheet().getSheetName(), row.getRowNum(), cellIndex), cell);
                rowData.addCellData(cellData);
            }else{
                rowData.addCellData(null);
            }
        }
        return rowData;
    }

    public Row getRow() {
        return row;
    }
}
