package org.jxls.common;

import java.util.HashMap;
import java.util.Map;

/**
 * Map bean context (Jxls internal class)
 * 
 * @author Leonid Vysochyn
 */
public class Context {
    protected Map<String, Object> varMap = new HashMap<String, Object>();
    private boolean formulaProcessingRequired = true;
    private boolean ignoreSourceCellStyle = false;
    private Map<String, String> cellStyleMap;
    
    public Context() {
    }

    public Context(Map<String, Object> varMap) {
        this.varMap = varMap;
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

    public boolean isFormulaProcessingRequired() {
        return formulaProcessingRequired;
    }

    public void setFormulaProcessingRequired(boolean formulaProcessingRequired) {
        this.formulaProcessingRequired = formulaProcessingRequired;
    }

    public boolean isIgnoreSourceCellStyle() {
        return ignoreSourceCellStyle;
    }

    public void setIgnoreSourceCellStyle(boolean ignoreSourceCellStyle) {
        this.ignoreSourceCellStyle = ignoreSourceCellStyle;
    }

    public Map<String, String> getCellStyleMap() {
        return cellStyleMap;
    }

    public void setCellStyleMap(Map<String, String> cellStyleMap) {
        this.cellStyleMap = cellStyleMap;
    }
}
