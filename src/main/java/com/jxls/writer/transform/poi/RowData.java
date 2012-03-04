package com.jxls.writer.transform.poi;

import com.jxls.writer.common.CellData;
import com.jxls.writer.common.CellRef;
import org.apache.poi.ss.usermodel.Row;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * @author Leonid Vysochyn
 *         Date: 2/1/12 2:01 PM
 */
public class RowData implements Iterable<CellData>{
    short height;
    List<CellData> cellDataList = new ArrayList<CellData>();
    Row row;

    public static RowData createRowData(Row row){
        if( row == null ) return null;
        RowData rowData = new RowData();
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
    
    public int getNumberOfCells(){
        return cellDataList.size();
    }
    
    public CellData getCellData(int col){
        if( col < cellDataList.size() ) return cellDataList.get(col);
        else return null;
    }
    
    public void addCellData(CellData cellData){
        cellDataList.add(cellData);
    }

    public short getHeight() {
        return height;
    }

    public Iterator<CellData> iterator() {
        return cellDataList.iterator();
    }

    public Row getRow() {
        return row;
    }
}
