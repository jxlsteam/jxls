package com.jxls.writer.builder.xls;

import com.jxls.writer.area.Area;
import com.jxls.writer.command.AbstractCommand;
import com.jxls.writer.command.Command;
import com.jxls.writer.common.CellRef;
import com.jxls.writer.common.Context;
import com.jxls.writer.common.Size;

import java.util.List;

/**
 * @author Leonid Vysochyn
 */
public class AreaCommand extends AbstractCommand {
    public String getName() {
        return "area";
    }

    public Size applyAt(CellRef cellRef, Context context) {
        return Size.ZERO_SIZE;
    }

}
