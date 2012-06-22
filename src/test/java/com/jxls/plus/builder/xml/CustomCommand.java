package com.jxls.plus.builder.xml;

import com.jxls.plus.command.AbstractCommand;
import com.jxls.plus.common.CellRef;
import com.jxls.plus.common.Context;
import com.jxls.plus.common.Size;

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
