package com.jxls.plus.builder.xml;

import ch.qos.logback.core.joran.action.Action;
import ch.qos.logback.core.joran.spi.ActionException;
import ch.qos.logback.core.joran.spi.InterpretationContext;
import com.jxls.plus.area.Area;
import com.jxls.plus.common.AreaRef;
import com.jxls.plus.command.Command;
import com.jxls.plus.command.IfCommand;
import org.xml.sax.Attributes;

/**
 * Builds {@link IfCommand} from XML
 * @author Leonid Vysochyn
 *         Date: 2/14/12 1:23 PM
 */
public class IfAction extends Action {
    public static final String REF_ATTR = "ref";
    public static final String CONDITION_ATTR = "condition";

    @Override
    public void begin(InterpretationContext ic, String name, Attributes attributes) throws ActionException {
        String ref = attributes.getValue(REF_ATTR);
        String condition = attributes.getValue(CONDITION_ATTR);
        if( condition == null || condition.length() == 0 ){
            String errMsg = "'condition' attribute of 'each' tag is empty";
            ic.addError(errMsg);
            throw new IllegalArgumentException(errMsg);
        }
        if( ref == null || ref.length() == 0 ){
            String errMsg = "'ref' attribute of 'each' tag is empty";
            ic.addError(errMsg);
        }
        Command command = new IfCommand(condition);
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
