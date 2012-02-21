package com.jxls.writer.builder.xml;

import com.jxls.writer.command.AbstractCommand;
import com.jxls.writer.common.CellRef;
import com.jxls.writer.common.Context;
import com.jxls.writer.common.Size;

/**
 * @author Leonid Vysochyn
 *         Date: 2/21/12 4:45 PM
 */
public class CustomCommand extends AbstractCommand{
    String attr;

    public String getName() {
        return "custom";
    }

    public Size applyAt(CellRef cellRef, Context context) {
        return Size.ZERO_SIZE;
    }

    public String getAttr() {
        return attr;
    }

    public void setAttr(String attr) {
        this.attr = attr;
    }
}
