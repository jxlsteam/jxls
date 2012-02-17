package com.jxls.writer.command;

import com.jxls.writer.common.CellRef;
import com.jxls.writer.common.Context;
import com.jxls.writer.common.Size;

/**
 * @author Leonid Vysochyn
 *         Date: 2/17/12 3:02 PM
 */
public class EachCellCommand extends AbstractCommand {

    String var;
    String items;
    CellRefGenerator cellRefGenerator;

    public String getName() {
        return "EachSheet";
    }

    public Size applyAt(CellRef cellRef, Context context) {
        return null;
    }
}
