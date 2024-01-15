package org.jxls.common;

/**
 * Jxls public context interface
 */
public interface PublicContext {

    Object getVar(String name);

    Object getRunVar(String name);

    void putVar(String name, Object value);

    void removeVar(String name);
    
    boolean containsVar(String name);

    Object evaluate(String expression);
    
    boolean isConditionTrue(String condition);
}
