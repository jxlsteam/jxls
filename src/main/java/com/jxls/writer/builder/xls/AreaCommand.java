package com.jxls.writer.builder.xls;

import com.jxls.writer.command.AbstractCommand;
import com.jxls.writer.common.CellRef;
import com.jxls.writer.common.Context;
import com.jxls.writer.common.Size;

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
