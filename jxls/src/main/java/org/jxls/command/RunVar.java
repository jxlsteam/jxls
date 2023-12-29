package org.jxls.command;

import org.jxls.common.PublicContext;

/**
 * Running variable handling
 */
public class RunVar implements AutoCloseable {
    private final String varName1;
    private final String varName2;
    private final Object oldValue1;
    private final Object oldValue2;
    private final PublicContext context;

    public RunVar(String varName, PublicContext context) {
        this(varName, null, context);
    }

    public RunVar(String varName1, String varName2, PublicContext context) {
        if (varName1 == null) {
            throw new IllegalArgumentException("varName must not be null");
        }
        this.varName1 = varName1;
        this.varName2 = varName2;
        this.context = context;
        oldValue1 = getRunVar(varName1, context);
        oldValue2 = varName2 == null ? null : getRunVar(varName2, context);
    }

    /**
     * Set value of running var
     * @param value -
     */
    public void put(Object value) {
        context.putVar(varName1, value);
    }

    /**
     * Set values of running vars
     * @param value1 -
     * @param value2 -
     */
    public void put(Object value1, Object value2) {
        context.putVar(varName1, value1);
        if (varName2 != null) {
            context.putVar(varName2, value2);
        }
    }

    @Override
    public void close() {
        // restore var values
        if (oldValue1 == null) {
            context.removeVar(varName1);
        } else {
            context.putVar(varName1, oldValue1);
        }
        if (varName2 != null) {
            if (oldValue2 == null) {
                context.removeVar(varName2);
            } else {
                context.putVar(varName2, oldValue2);
            }
        }
    }

    public static Object getRunVar(String varName, PublicContext context) {
        if (varName != null && context != null) {
            if (context.containsVar(varName)) {
                return context.getRunVar(varName);
            }
        }
        return null;
    }
}
