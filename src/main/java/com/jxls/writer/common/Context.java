package com.jxls.writer.common;

import java.util.HashMap;
import java.util.Map;

/**
 * Map bean context
 * Date: Nov 2, 2009
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

    public Map<String, Object> toMap(){
        return varMap;
    }
    
    public Object getVar(String name){
        if( varMap.containsKey(name) ) return varMap.get(name);
        else return null;
    }

    public void putVar(String name, Object value) {
        varMap.put(name, value);
    }

    public void removeVar(String var) {
        varMap.remove(var);
    }

    @Override
    public String toString() {
        return "Context" +
                varMap;
    }
}
