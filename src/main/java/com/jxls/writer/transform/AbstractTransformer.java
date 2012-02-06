package com.jxls.writer.transform;

import com.jxls.writer.CellData;
import com.jxls.writer.Pos;
import com.jxls.writer.command.Context;

import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

/**
 * @author Leonid Vysochyn
 *         Date: 2/6/12 6:33 PM
 */
public abstract class AbstractTransformer implements Transformer{

    protected CellData[][][] cellData;
    boolean ignoreColumnProps = false;
    boolean ignoreRowProps = false;

    public List<Pos> getTargetPos(Pos pos) {
        CellData cellData = getCellData(pos);
        if( cellData != null ){
            return cellData.getTargetPos();
        }else{
            return new ArrayList<Pos>();
        }
    }

    public CellData getCellData(Pos pos){
        return cellData[pos.getSheet()][pos.getRow()][pos.getCol()];
    }

    public boolean isIgnoreColumnProps() {
        return ignoreColumnProps;
    }

    public void setIgnoreColumnProps(boolean ignoreColumnProps) {
        this.ignoreColumnProps = ignoreColumnProps;
    }

    public boolean isIgnoreRowProps() {
        return ignoreRowProps;
    }

    public void setIgnoreRowProps(boolean ignoreRowProps) {
        this.ignoreRowProps = ignoreRowProps;
    }

    public Set<CellData> getFormulaCells() {
        Set<CellData> formulaCells = new HashSet<CellData>();
        for(int sheetInd = 0; sheetInd < cellData.length; sheetInd++){
            for(int rowInd = 0; rowInd < cellData[sheetInd].length; rowInd++){
                if(cellData[sheetInd][rowInd] != null ){
                    for(int colInd = 0; colInd < cellData[sheetInd][rowInd].length; colInd++){
                        if(cellData[sheetInd][rowInd][colInd]!= null && cellData[sheetInd][rowInd][colInd].isFormulaCell() ){
                            formulaCells.add(cellData[sheetInd][rowInd][colInd]);
                        }
                    }
                }
            }
        }
        return formulaCells;
    }
}
