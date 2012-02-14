package com.jxls.writer.command;

import com.jxls.writer.*;

import java.util.ArrayList;
import java.util.List;

/**
 * @author Leonid Vysochyn
 *         Date: 21.03.2009
 */
public abstract class AbstractCommand implements Command {
    Size initialSize;
    List<Area> areaList = new ArrayList<Area>();

    public Size getInitialSize() {
        return initialSize;
    }

    public void addArea(Area area) {
        areaList.add(area);
    }

    public List<Area> getAreaList() {
        return areaList;
    }
}
