package com.jxls.plus.command;

import com.jxls.plus.area.Area;
import com.jxls.plus.transform.Transformer;

import java.util.ArrayList;
import java.util.List;

/**
 * Implements basic command methods and is a convenient base class for other commands
 * @author Leonid Vysochyn
 *         Date: 21.03.2009
 */
public abstract class AbstractCommand implements Command {
    List<Area> areaList = new ArrayList<Area>();

    public Command addArea(Area area) {
        areaList.add(area);
        return this;
    }

    public void reset() {
        for (Area area : areaList) {
            area.reset();
        }
    }

    public List<Area> getAreaList() {
        return areaList;
    }

    protected Transformer getTransformer(){
        if( areaList.isEmpty() ) return null;
        return areaList.get(0).getTransformer();
    }
}
