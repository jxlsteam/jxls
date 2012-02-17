package com.jxls.writer.command;

import com.jxls.writer.common.CellRef;
import com.jxls.writer.common.Context;

/**
 * @author Leonid Vysochyn
 *         Date: 2/17/12 5:22 PM
 */
public interface CellRefGenerator {
    CellRef generateCellRef(int index, Context context);
}
