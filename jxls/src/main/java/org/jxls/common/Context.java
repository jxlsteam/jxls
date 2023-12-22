package org.jxls.common;

import java.util.HashMap;
import java.util.Map;

/**
 * Map bean context
 * 
 * @author Leonid Vysochyn
 */
public class Context {
    protected Map<String, Object> varMap = new HashMap<String, Object>();

    public Context() {
    }

    public Context(Map<String, Object> varMap) {
        for (Map.Entry<String, Object> entry : varMap.entrySet()) {
            this.varMap.put(entry.getKey(), entry.getValue());
        }
    }

    public Map<String, Object> toMap() {
        return varMap;
    }

    public Object getVar(String name) {
        return varMap.get(name);
    }

    public Object getRunVar(String name) {
        return getVar(name);
    }

    public void putVar(String name, Object value) {
        varMap.put(name, value);
    }

    public void removeVar(String var) {
        varMap.remove(var);
    }

    @Override
    public String toString() {
        return "Context" + varMap;
    }
}
