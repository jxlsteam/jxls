package org.jxls.builder.xml;

import org.jxls.area.Area;
import org.jxls.command.EachCommand;
import org.jxls.common.AreaRef;
import org.xml.sax.Attributes;

import ch.qos.logback.core.joran.action.Action;
import ch.qos.logback.core.joran.spi.ActionException;
import ch.qos.logback.core.joran.spi.InterpretationContext;

/**
 * Builds {@link org.jxls.command.EachCommand} from XML
 * 
 * @author Leonid Vysochyn
 */
public class EachAction extends Action {
    public static final String ITEMS_ATTR = "items";
    public static final String VAR_ATTR = "var";
    public static final String REF_ATTR = "ref";
    public static final String DIRECTION_ATTR = "dir";
    public static final String GROUP_BY_ATTR = "groupBy";
    public static final String GROUP_ORDER_ATTR = "groupOrder";
    public static final String ORDER_BY_ATTR = "orderBy";

    @Override
    public void begin(InterpretationContext ic, String name, Attributes attributes) throws ActionException {
        String items = attributes.getValue(ITEMS_ATTR);
        String var = attributes.getValue(VAR_ATTR);
        String ref = attributes.getValue(REF_ATTR);
        String dir = attributes.getValue(DIRECTION_ATTR);
        String groupBy = attributes.getValue(GROUP_BY_ATTR); // optional
        String groupOrder = attributes.getValue(GROUP_ORDER_ATTR); // optional
        String orderBy = attributes.getValue(ORDER_BY_ATTR); // optional
        EachCommand.Direction direction;
        if (items == null || items.length() == 0) {
            String errMsg = "'items' attribute of 'each' tag is empty";
            ic.addError(errMsg);
            throw new IllegalArgumentException(errMsg);
        }
        if (ref == null || ref.length() == 0) {
            String errMsg = "'ref' attribute of 'each' tag is empty";
            ic.addError(errMsg);
        }
        if (dir == null || dir.length() == 0) {
            direction = EachCommand.Direction.DOWN;
        } else {
            try {
                direction = EachCommand.Direction.valueOf(dir);
            } catch (IllegalArgumentException e) {
                String errMsg = "'dir' attribute [" + dir + "] cannot be converted to EachCommand.Direction instance";
                ic.addError(errMsg);
                throw new IllegalArgumentException(errMsg);
            }
        }
        EachCommand command = new EachCommand(var, items, direction);
        command.setGroupBy(groupBy);
        command.setGroupOrder(groupOrder);
        command.setOrderBy(orderBy);
        Object object = ic.peekObject();
        if (object instanceof Area) {
            Area area = (Area) object;
            area.addCommand(new AreaRef(ref), command);
        } else {
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
