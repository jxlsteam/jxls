package com.jxls.writer.builder.xml;

import ch.qos.logback.core.joran.action.Action;
import ch.qos.logback.core.joran.spi.ActionException;
import ch.qos.logback.core.joran.spi.InterpretationContext;
import com.jxls.writer.area.Area;
import com.jxls.writer.common.AreaRef;
import com.jxls.writer.command.Command;
import com.jxls.writer.command.EachCommand;
import org.xml.sax.Attributes;

/**
 * Builds {@link EachCommand} from XML
 * @author Leonid Vysochyn
 *         Date: 2/14/12
 */
public class EachAction extends Action {
    public static final String ITEMS_ATTR = "items";
    public static final String VAR_ATTR = "var";
    public static final String REF_ATTR = "ref";
    public static final String DIRECTION_ATTR = "dir";

    @Override
    public void begin(InterpretationContext ic, String name, Attributes attributes) throws ActionException {
        String items = attributes.getValue(ITEMS_ATTR);
        String var = attributes.getValue(VAR_ATTR);
        String ref = attributes.getValue(REF_ATTR);
        String dir = attributes.getValue(DIRECTION_ATTR);
        EachCommand.Direction direction;
        if( items == null || items.length() == 0 ){
            String errMsg = "'items' attribute of 'each' tag is empty";
            ic.addError(errMsg);
            throw new IllegalArgumentException(errMsg);
        }
        if( ref == null || ref.length() == 0 ){
            String errMsg = "'ref' attribute of 'each' tag is empty";
            ic.addError(errMsg);
        }
        if( dir == null || dir.length() == 0 ){
            direction = EachCommand.Direction.DOWN;
        }else {
            try{
                direction = EachCommand.Direction.valueOf(dir);
            }catch (IllegalArgumentException e){
                String errMsg = "'dir' attribute [" + dir + "] cannot be converted to EachCommand.Direction instance";
                ic.addError(errMsg);
                throw new IllegalArgumentException(errMsg);
            }
        }
        Command command = new EachCommand(var, items, direction);
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
