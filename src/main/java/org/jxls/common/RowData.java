package org.jxls.common;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * Represents an excel row data holder
 * @author Leonid Vysochyn
 */
public class RowData implements Iterable<CellData> {
    protected int height;
    protected List<CellData> cellDataList = new ArrayList<CellData>();

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

    public int getHeight() {
        return height;
    }

    public Iterator<CellData> iterator() {
        return cellDataList.iterator();
    }
}
