package com.jxls.writer.transform;

import com.jxls.writer.Cell;
import com.jxls.writer.CellData;
import com.jxls.writer.Pos;
import com.jxls.writer.command.Context;

import java.util.List;
import java.util.Set;

/**
 * @author Leonid Vysochyn
 *         Date: 1/23/12 1:24 PM
 */
public interface Transformer {
    void transform(Cell cell, Cell newCell, Context context);
    void updateFormulaCell(Cell cell, String formulaString);
    Set<CellData> getFormulaCells();
    CellData getCellData(int sheet, int row, int col);
    List<Pos> getTargetCells(int sheet, int row, int col);
}
