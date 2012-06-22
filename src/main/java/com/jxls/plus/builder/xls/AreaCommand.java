package com.jxls.plus.builder.xls;

import com.jxls.plus.common.Size;
import com.jxls.plus.command.AbstractCommand;
import com.jxls.plus.common.CellRef;
import com.jxls.plus.common.Context;

/**
 * A container area command used in {@link XlsCommentAreaBuilder} to enclose other commands
 * @author Leonid Vysochyn
 */
public class AreaCommand extends AbstractCommand {

    String clearCells;

    public String getName() {
        return "area";
    }

    public Size applyAt(CellRef cellRef, Context context) {
        return Size.ZERO_SIZE;
    }

    public String getClearCells() {
        return clearCells;
    }

    public void setClearCells(String clearCells) {
        this.clearCells = clearCells;
    }
}
