package org.jxls.command;

import org.jxls.area.Area;
import org.jxls.expression.ExpressionEvaluator;
import org.jxls.expression.JexlExpressionEvaluator;
import org.jxls.transform.Transformer;

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

    protected ExpressionEvaluator getExpressionEvaluator(){
        if(getTransformer() != null ){
            return getTransformer().getExpressionEvaluator();
        }else{
            return new JexlExpressionEvaluator();
        }
    }
}
