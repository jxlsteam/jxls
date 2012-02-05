package com.jxls.writer.transform;

import com.jxls.writer.Pos;
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
    void transform(Pos pos, Pos newPos, Context context);
    void updateFormulaCell(Pos pos, String formulaString);
    Set<CellData> getFormulaCells();
    CellData getCellData(int sheet, int row, int col);
    Set<Pos> getTargetPos(int sheet, int row, int col);
}
