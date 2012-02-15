package com.jxls.writer.common;

import java.util.HashMap;
import java.util.Map;

/**
 * Date: Nov 2, 2009
 *
 * @author Leonid Vysochyn
 */
public class Context {
    Map<String, Object> varMap = new HashMap<String, Object>();

    public Map<String, Object> toMap(){
        return varMap;
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
