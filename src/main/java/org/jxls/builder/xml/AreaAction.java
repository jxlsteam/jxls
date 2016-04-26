package org.jxls.builder.xml;

import ch.qos.logback.core.joran.action.Action;
import ch.qos.logback.core.joran.spi.ActionException;
import ch.qos.logback.core.joran.spi.InterpretationContext;
import org.jxls.area.Area;
import org.jxls.area.XlsArea;
import org.jxls.command.Command;
import org.jxls.transform.Transformer;
import org.xml.sax.Attributes;

import java.util.ArrayList;
import java.util.List;

/**
 * Builds {@link org.jxls.builder.xls.AreaCommand} from XML
 * @author Leonid Vysochyn
 *         Date: 2/14/12
 */
class AreaAction extends Action {
    private static final String REF_ATTR = "ref";
    private List<Area> areaList = new ArrayList<Area>();
    private Transformer transformer;

    AreaAction(Transformer transformer) {
        this.transformer = transformer;
    }

    @Override
    public void begin(InterpretationContext ic, String name, Attributes attributes) throws ActionException {
        String ref = attributes.getValue( REF_ATTR );
        Area area = new XlsArea(ref, transformer);
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
