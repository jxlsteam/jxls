package com.jxls.writer.command;

import com.jxls.writer.common.CellRef;
import com.jxls.writer.common.Context;

/**
 * Defines generic method for generating new cell references during {@link EachCommand} command iteration
 * @author Leonid Vysochyn
 *         Date: 2/17/12
 */
public interface CellRefGenerator {
    /**
     *
     * @param index current iteration index
     * @param context current context
     * @return target cell reference for this index
     */
    CellRef generateCellRef(int index, Context context);
}
