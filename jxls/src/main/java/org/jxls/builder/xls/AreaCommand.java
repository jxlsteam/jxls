package org.jxls.builder.xls;

import org.jxls.command.AbstractCommand;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.Size;

/**
 * A container area command used in {@link XlsCommentAreaBuilder} to enclose other commands
 * 
 * @author Leonid Vysochyn
 */
public class AreaCommand extends AbstractCommand {
    public static final String COMMAND_NAME = "area";

    private String clearCells;

    @Override
    public String getName() {
        return COMMAND_NAME;
    }

    @Override
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
