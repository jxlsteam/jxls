package org.jxls.builder.xml;

import ch.qos.logback.core.joran.action.Action;
import ch.qos.logback.core.joran.spi.ActionException;
import ch.qos.logback.core.joran.spi.InterpretationContext;
import org.jxls.area.Area;
import org.jxls.command.Command;
import org.jxls.command.GridCommand;
import org.jxls.common.AreaRef;
import org.xml.sax.Attributes;

/**
 * Defines Grid command via xml
 *
 * @author Leonid Vysochyn
 */
public class GridAction extends Action {
    private static final String HEADER_ATTR = "headers";
    private static final String DATA_ATTR = "data";
    private static final String REF_ATTR = "ref";

    @Override
    public void begin(InterpretationContext ic, String s, Attributes attributes) throws ActionException {
        String headers = attributes.getValue(HEADER_ATTR);
        String data = attributes.getValue(DATA_ATTR);
        String ref = attributes.getValue(REF_ATTR);
        if (headers == null || headers.length() == 0) {
            String errMsg = "'headers' attribute of 'grid' tag is empty";
            ic.addError(errMsg);
            throw new IllegalArgumentException(errMsg);
        }
        if (ref == null || ref.length() == 0) {
            String errMsg = "'ref' attribute of 'grid' tag is empty";
            ic.addError(errMsg);
        }
        Command command = new GridCommand(headers, data);
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
    public void end(InterpretationContext ic, String s) throws ActionException {
        ic.popObject();
    }
}
