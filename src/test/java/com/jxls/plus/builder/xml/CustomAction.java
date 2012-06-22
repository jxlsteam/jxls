package com.jxls.plus.builder.xml;

import ch.qos.logback.core.joran.action.Action;
import ch.qos.logback.core.joran.spi.ActionException;
import ch.qos.logback.core.joran.spi.InterpretationContext;
import com.jxls.plus.area.Area;
import com.jxls.plus.common.AreaRef;
import org.xml.sax.Attributes;

/**
 * @author Leonid Vysochyn
 *         Date: 2/21/12 5:27 PM
 */
public class CustomAction extends Action {
    public static final String ATTR = "attr";
    public static final String REF_ATTR = "ref";

    @Override
    public void begin(InterpretationContext ic, String name, Attributes attributes) throws ActionException {
        String attr = attributes.getValue(ATTR);
        String ref = attributes.getValue(REF_ATTR);
        if( ref == null || ref.length() == 0 ){
            String errMsg = "'ref' attribute of 'custom' tag is empty";
            ic.addError(errMsg);
        }
        CustomCommand command = new CustomCommand();
        command.setAttr(attr);
        Object object = ic.peekObject();
        if( object instanceof Area){
            Area area = (Area) object;
            area.addCommand(new AreaRef(ref), command);
        }else{
            String errMsg = "Object [" + object + "] currently at the top of the stack is not an Area";
            ic.addError(errMsg);
            throw new IllegalArgumentException(errMsg);
        }
        ic.pushObject(command);
    }

    @Override
    public void end(InterpretationContext ic, String name) throws ActionException {
        ic.popObject();
    }
}