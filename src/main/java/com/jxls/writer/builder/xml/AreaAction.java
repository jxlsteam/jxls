package com.jxls.writer.builder.xml;

import ch.qos.logback.core.joran.action.Action;
import ch.qos.logback.core.joran.spi.ActionException;
import ch.qos.logback.core.joran.spi.InterpretationContext;
import com.jxls.writer.command.Area;
import com.jxls.writer.command.Command;
import com.jxls.writer.command.XlsArea;
import com.jxls.writer.transform.Transformer;
import org.xml.sax.Attributes;

import java.util.ArrayList;
import java.util.List;

/**
 * @author Leonid Vysochyn
 *         Date: 2/14/12 1:24 PM
 */
public class AreaAction extends Action {
    public static final String REF_ATTR = "ref";
    List<Area> areaList = new ArrayList<Area>();
    Transformer transformer;

    public AreaAction(Transformer transformer) {
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
