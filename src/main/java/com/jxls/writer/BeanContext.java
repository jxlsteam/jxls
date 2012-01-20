package com.jxls.writer;

import java.util.HashMap;
import java.util.Map;

/**
 * Date: Apr 13, 2009
 *
 * @author Leonid Vysochyn
 */
public class BeanContext {
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
}
