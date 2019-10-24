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
    private Config config = new Config();

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

    public void putVar(String name, Object value) {
        varMap.put(name, value);
    }

    public void removeVar(String var) {
        varMap.remove(var);
    }

    public Config getConfig() {
        return config;
    }

    @Override
    public String toString() {
        return "Context" + varMap;
    }

    /**
     * Special config class to use in Area processing
     */
    public class Config {
        private boolean ignoreSourceCellStyle = false;
        private Map<String, String> cellStyleMap = new HashMap<>();
        private boolean isFormulaProcessingRequired = true;

        public boolean isFormulaProcessingRequired() {
            return isFormulaProcessingRequired;
        }

        public void setIsFormulaProcessingRequired(boolean isFormulaProcessingRequired) {
            this.isFormulaProcessingRequired = isFormulaProcessingRequired;
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
}
