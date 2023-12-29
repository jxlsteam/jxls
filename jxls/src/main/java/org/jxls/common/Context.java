package org.jxls.common;

import java.util.HashMap;
import java.util.Map;

import org.jxls.expression.ExpressionEvaluator;
import org.jxls.expression.ExpressionEvaluatorFactoryJexlImpl;
import org.jxls.transform.ExpressionEvaluatorContext;

/**
 * Jxls context (Jxls internal class)
 * 
 * @author Leonid Vysochyn
 */
public class Context {
    private final ExpressionEvaluatorContext expressionEvaluatorContext;
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

    public Context(ExpressionEvaluatorContext expressionEvaluatorContext, Map<String, Object> varMap) {
        this.expressionEvaluatorContext = expressionEvaluatorContext == null ? new ExpressionEvaluatorContext(new ExpressionEvaluatorFactoryJexlImpl(), null, null) : expressionEvaluatorContext;
        this.varMap = varMap;
    }
    
    public Object evaluate(String expression) {
        return getExpressionEvaluator(expression).evaluate(varMap);
    }

    public ExpressionEvaluator getExpressionEvaluator(String expression) {
        return expressionEvaluatorContext.getExpressionEvaluator(expression);
    }

    /**
     * INTERNAL
     * @param rawExpression e.g. "${e.name}"
     * @return EvaluationResult
     */
    public EvaluationResult _evaluateRawExpression(String rawExpression) {
        return expressionEvaluatorContext.evaluateRawExpression(rawExpression, varMap);
    }
    
    public boolean isConditionTrue(String condition) {
        return getExpressionEvaluator(condition).isConditionTrue(varMap);
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
