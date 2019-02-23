package org.jxls.common;

import org.jxls.transform.Transformer;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * Represents an excel row data holder
 * @author Leonid Vysochyn
 */
public class RowData implements Iterable<CellData> {
    protected int height;
    private List<CellData> cellDataList = new ArrayList<CellData>();
    private Transformer transformer;

    public Transformer getTransformer() {
        return transformer;
    }

    public void setTransformer(Transformer transformer) {
        this.transformer = transformer;
    }

    public int getNumberOfCells(){
        return cellDataList.size();
    }

    public CellData getCellData(int col){
        if( col < cellDataList.size() ) return cellDataList.get(col);
        else return null;
    }

    protected void addCellData(CellData cellData){
        cellDataList.add(cellData);
    }

    public int getHeight() {
        return height;
    }

    public Iterator<CellData> iterator() {
        return cellDataList.iterator();
    }
}
