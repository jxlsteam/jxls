package org.jxls.common;

import java.util.HashMap;
import java.util.Map;

import org.jxls.expression.ExpressionEvaluatorFactoryJexlImpl;
import org.jxls.transform.TransformationConfig;

/**
 * Jxls context (Jxls internal class)
 * 
 * @author Leonid Vysochyn
 */
public class Context {
    private final TransformationConfig transformationConfig;
    private final Map<String, Object> varMap;
    private boolean formulaProcessingRequired = true;
    private boolean ignoreSourceCellStyle = false;
    private Map<String, String> cellStyleMap;
    
    /**
     * Should only be used for Jxls internal testcases
     */
    public Context() {
        this(null, new HashMap<String, Object>());
    }

    public Context(TransformationConfig transformationConfig, Map<String, Object> varMap) {
        this.transformationConfig = transformationConfig == null ? new TransformationConfig(new ExpressionEvaluatorFactoryJexlImpl(), null, null) : transformationConfig;
        this.varMap = varMap;
    }
    
    public TransformationConfig getTransformationConfig() {
        return transformationConfig;
    }
    
    public boolean isConditionTrue(String condition) {
        return transformationConfig.getExpressionEvaluator().isConditionTrue(condition, toMap());
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
