package com.jxls.writer.transform;

import com.jxls.writer.Cell;
import com.jxls.writer.CellData;
import com.jxls.writer.command.Context;

import java.util.List;

/**
 * @author Leonid Vysochyn
 *         Date: 1/23/12 1:24 PM
 */
public interface Transformer {
    void transform(Cell cell, Cell newCell, Context context);
    void updateFormulaCell(Cell cell, String formulaString);
    List<CellData> getFormulaCells();
}
