package com.jxls.writer.command;

import com.jxls.writer.area.Area;
import com.jxls.writer.common.CellRef;
import com.jxls.writer.common.Context;
import com.jxls.writer.common.Size;

import java.util.List;

/**
 * Date: Mar 13, 2009
 *
 * @author Leonid Vysochyn
 */
public interface Command {

    String getName();
    List<Area> getAreaList();
    Command addArea(Area area);

    Size applyAt(CellRef cellRef, Context context);

}
