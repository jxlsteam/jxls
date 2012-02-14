package com.jxls.writer.command;

import com.jxls.writer.*;

import java.util.List;

/**
 * Date: Mar 13, 2009
 *
 * @author Leonid Vysochyn
 */
public interface Command {

    String getName();
    List<Area> getAreaList();
    void addArea(Area area);

    Size applyAt(CellRef cellRef, Context context);

}
