package org.jxls.common;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.jxls.transform.Transformer;

/**
 * Represents an Excel row data holder
 * 
 * @author Leonid Vysochyn
 */
public class RowData implements Iterable<CellData> {
    protected int height;
    private final List<CellData> cellDataList = new ArrayList<CellData>();
    private Transformer transformer;

    public Transformer getTransformer() {
        return transformer;
    }

    public void setTransformer(Transformer transformer) {
        this.transformer = transformer;
    }

    public int getNumberOfCells() {
        return cellDataList.size();
    }

    public CellData getCellData(int col) {
        if (col < cellDataList.size()) {
            return cellDataList.get(col);
        } else {
            return null;
        }
    }

    protected void addCellData(CellData cellData){
        cellDataList.add(cellData);
    }

    public int getHeight() {
        return height;
    }

    @Override
    public Iterator<CellData> iterator() {
        return cellDataList.iterator();
    }
}
