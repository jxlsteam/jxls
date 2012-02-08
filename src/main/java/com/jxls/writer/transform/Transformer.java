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
    void setFormula(Pos pos, String formulaString);
    Set<CellData> getFormulaCells();
    CellData getCellData(Pos pos);
    List<Pos> getTargetPos(Pos pos);
    void resetTargetCells();
}
