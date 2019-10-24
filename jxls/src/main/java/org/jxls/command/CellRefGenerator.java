package org.jxls.command;

import org.jxls.common.CellRef;
import org.jxls.common.Context;

/**
 * Defines generic method for generating new cell references during {@link EachCommand} command iteration
 * 
 * @author Leonid Vysochyn
 */
public interface CellRefGenerator {
    
    /**
     * @param index current iteration index
     * @param context current context
     * @return target cell reference for this index
     */
    CellRef generateCellRef(int index, Context context);
}
