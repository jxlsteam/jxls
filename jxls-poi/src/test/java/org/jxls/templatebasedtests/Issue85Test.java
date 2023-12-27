package org.jxls.templatebasedtests;

import java.util.HashMap;
import java.util.Map;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.common.Context;
import org.jxls.entity.Employee;

/**
 * given scenario: You have a self implemented Context and Map. The map doesn't like unknown keys.
 * 
 * When accessing the map, you don't know whether a missing key is really missing or whether it is just the run var.
 * EachCommand does save the run var before the loop and sets it back after the loop.
 */
public class Issue85Test {

    @Test
    public void test() throws Exception {
        JxlsTester.xlsx(getClass()).processTemplate(new ContextWithMyMap(new MyMap()));
    }
    
    private static class MyMap extends HashMap<String, Object> {
        
        public MyMap() {
            put("title", "the title");
            put("employees", Employee.generateSampleEmployeeData());
        }
        
        @Override
        public Object get(Object key) {
            if ("jx:isFormulaProcessingRequired".equals(key)) return null; // TODO
            if (!containsKey(key)) {
                throw new IllegalArgumentException("Map does not contain key: " + key);
            }
            return super.get(key);
        }
    }
    
    private static class ContextWithMyMap extends Context {
        private final Map<String, Object> myMap;
        
        public ContextWithMyMap(Map<String, Object> myMap) {
            this.myMap = myMap;
        }
        
        @Override
        public Map<String, Object> toMap() {
            return myMap;
        }
        
        @Override
        public Object getVar(String name) {
            return myMap.get(name);
        }
        
        @Override
        public Object getRunVar(String name) {
            try {
                return super.getRunVar(name);
            } catch (Exception ignore) {
                return null;
            }
        }
        
        @Override
        public void putVar(String name, Object value) {
            myMap.put(name, value);
        }
        
        @Override
        public void removeVar(String key) {
            myMap.remove(key);
        }
    }
}
