package org.jxls.command;

import java.util.ArrayList;
import java.util.List;

import org.jxls.area.Area;
import org.jxls.command.EachCommand.Direction;
import org.jxls.common.Context;
import org.jxls.util.UtilWrapper;

/**
 * Filter collection
 */
public class CollectionFilter extends CollectionProcessor {
    private final List<Object> resultList = new ArrayList<>();
    
    public CollectionFilter(Context context, UtilWrapper util, String varIndex, Direction direction, String select,
            String groupBy, String multisheet, CellRefGenerator cellRefGenerator, Area area) {
        super(context, util, varIndex, direction, select, groupBy, multisheet, cellRefGenerator, area);
    }

    @Override
    protected boolean processItem(Object item) {
        resultList.add(item);
        return false;
    }

    public List<Object> getFilteredCollection() {
        return resultList;
    }
}
