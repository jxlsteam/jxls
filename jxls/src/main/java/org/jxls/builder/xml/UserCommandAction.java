package org.jxls.builder.xml;

import ch.qos.logback.core.joran.action.Action;
import ch.qos.logback.core.joran.spi.ActionException;
import ch.qos.logback.core.joran.spi.InterpretationContext;
import ch.qos.logback.core.util.OptionHelper;
import org.jxls.area.Area;
import org.jxls.common.AreaRef;
import org.jxls.util.Util;
import org.jxls.command.Command;
import org.xml.sax.Attributes;

/**
 * Builds custom user command from XML
 * 
 * @author Leonid Vysochyn
 * @since 2/21/12
 */
public class UserCommandAction extends Action {
    public static final String ATTR = "attr";
    public static final String REF_ATTR = "ref";
    public static final String COMMAND_CLASS_ATTR = "commandClass";

    @Override
    public void begin(InterpretationContext ic, String name, Attributes attributes) throws ActionException {
        String ref = attributes.getValue(REF_ATTR);
        String commandClassName = attributes.getValue(COMMAND_CLASS_ATTR);
        if (commandClassName == null || commandClassName.trim().length() == 0) {
            String errMsg = "Required actionClass attribute is not specified for userCommand";
            ic.addError(errMsg);
            throw new IllegalArgumentException(errMsg);
        }
        Command command;
        try {
            command = (Command) OptionHelper.instantiateByClassName(commandClassName, Command.class, context);
        } catch (Exception e) {
            addError("Could not instantiate class [" + commandClassName + "]", e);
            throw new IllegalStateException(e);
        }
        try {
            initPropertiesFromAttributes(command, attributes);
        } catch (Exception e) {
            addWarn("Could not set an attribute");
        }
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

    private void initPropertiesFromAttributes(Object obj, Attributes attributes) {
        int attrLength = attributes.getLength();
        for (int i = 0; i < attrLength; i++) {
            try {
                Util.setObjectProperty(obj, attributes.getLocalName(i), attributes.getValue(i));
            } catch (Exception e) {
                addWarn("Could not set an attribute attr=" + attributes.getLocalName(i) + ", value=" + attributes.getValue(i));
            }
        }
    }

    @Override
    public void end(InterpretationContext ic, String name) throws ActionException {
        ic.popObject();
    }
}