package com.jxls.plus.builder.xml;

import ch.qos.logback.core.joran.action.Action;
import ch.qos.logback.core.joran.spi.ActionException;
import ch.qos.logback.core.joran.spi.InterpretationContext;
import com.jxls.plus.area.Area;
import com.jxls.plus.area.XlsArea;
import com.jxls.plus.command.Command;
import com.jxls.plus.transform.Transformer;
import org.xml.sax.Attributes;

import java.util.ArrayList;
import java.util.List;

/**
 * Builds {@link AreaCommand} from XML
 * @author Leonid Vysochyn
 *         Date: 2/14/12
 */
public class AreaAction extends Action {
    public static final String REF_ATTR = "ref";
    public static final String CLEAR_CELLS_ATTR = "clearCells";
    List<Area> areaList = new ArrayList<Area>();
    Transformer transformer;

    public AreaAction(Transformer transformer) {
        this.transformer = transformer;
    }

    @Override
    public void begin(InterpretationContext ic, String name, Attributes attributes) throws ActionException {
        String ref = attributes.getValue( REF_ATTR );
        String clearCellsFlagStr = attributes.getValue(CLEAR_CELLS_ATTR);
        boolean clearCellsFlag = false;
        Area area = new XlsArea(ref, transformer);
        if( clearCellsFlagStr != null && clearCellsFlagStr.equalsIgnoreCase("true")){
            clearCellsFlag = true;
        }
        ((XlsArea)area).setClearCellsBeforeApply(clearCellsFlag);
        if(!ic.isEmpty()){
            Object object = ic.peekObject();
            if( object instanceof Command){
                Command command = (Command) object;
                command.addArea(area);
            }else{
                String errMsg = "Object [" + object + "] currently at the top of the stack is not a Command";
                ic.addError(errMsg);
                throw new IllegalArgumentException(errMsg);
            }
        }
        ic.pushObject(area);
    }

    @Override
    public void end(InterpretationContext ic, String name) throws ActionException {
        Area area = (Area) ic.popObject();
        if(ic.isEmpty()){
            areaList.add(area);
        }
    }

    public List<Area> getAreaList() {
        return areaList;
    }
}
