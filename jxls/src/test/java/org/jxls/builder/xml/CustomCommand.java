package org.jxls.builder.xml;

import org.jxls.command.AbstractCommand;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.Size;

/**
 * @author Leonid Vysochyn
 */
public class CustomCommand extends AbstractCommand {
    private String attr;

    @Override
    public String getName() {
        return "custom";
    }

    @Override
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
