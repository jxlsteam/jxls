package com.jxls.writer.command;

import com.jxls.writer.common.Size;
import com.jxls.writer.common.AreaRef;
import com.jxls.writer.common.CellRef;
import com.jxls.writer.transform.Transformer;

import java.util.List;

/**
 * @author Leonid Vysochyn
 *         Date: 1/18/12 5:20 PM
 */
public interface Area{
    Size applyAt(CellRef cellRef, Context context);

    CellRef getStartCellRef();
    Size getSize();

    List<CommandData> getCommandDataList();
    void addCommand(AreaRef ref, Command command);

    Transformer getTransformer();
    void processFormulas();
}
