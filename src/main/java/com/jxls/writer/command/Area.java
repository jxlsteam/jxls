package com.jxls.writer.command;

import com.jxls.writer.CellRef;
import com.jxls.writer.Size;
import com.jxls.writer.transform.Transformer;

/**
 * @author Leonid Vysochyn
 *         Date: 1/18/12 5:20 PM
 */
public interface Area extends Command{
    CellRef getStartCellRef();
    Transformer getTransformer();
    void processFormulas();

    Size getInitialSize();
}
