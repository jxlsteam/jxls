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
public class ContextImpl implements Context {
    private final ExpressionEvaluatorContext expressionEvaluatorContext;
    private final Map<String, Object> varMap;
    private boolean formulaProcessingRequired = true;
    private boolean ignoreSourceCellStyle = false;
    private Map<String, String> cellStyleMap;
    
    /**
     * Should only be used for Jxls internal testcases
     */
    public ContextImpl() {
        this(null, new HashMap<String, Object>());
    }

    public ContextImpl(ExpressionEvaluatorContext expressionEvaluatorContext, Map<String, Object> varMap) {
        this.expressionEvaluatorContext = expressionEvaluatorContext == null ? new ExpressionEvaluatorContext(new ExpressionEvaluatorFactoryJexlImpl(), null, null) : expressionEvaluatorContext;
        this.varMap = varMap;
    }
    
    @Override
    public Object evaluate(String expression) {
        return getExpressionEvaluator(expression).evaluate(varMap);
    }

    @Override
    public ExpressionEvaluator getExpressionEvaluator(String expression) {
        return expressionEvaluatorContext.getExpressionEvaluator(expression);
    }

    /**
     * INTERNAL
     * @param rawExpression e.g. "${e.name}"
     * @return EvaluationResult
     */
    @Override
    public EvaluationResult _evaluateRawExpression(String rawExpression) {
        return expressionEvaluatorContext.evaluateRawExpression(rawExpression, varMap);
    }
    
    @Override
    public boolean isConditionTrue(String condition) {
        return getExpressionEvaluator(condition).isConditionTrue(varMap);
    }

    @Override
    public Map<String, Object> toMap() {
        return varMap;
    }

    @Override
    public Object getVar(String name) {
        return varMap.get(name);
    }

    @Override
    public Object getRunVar(String name) {
        return getVar(name);
    }

    @Override
    public void putVar(String name, Object value) {
        varMap.put(name, value);
    }

    @Override
    public void removeVar(String var) {
        varMap.remove(var);
    }
    
    @Override
    public boolean containsVar(String name) {
        return varMap.containsKey(name);
    }

    @Override
    public String toString() {
        return "Context" + varMap;
    }

    @Override
    public boolean isFormulaProcessingRequired() {
        return formulaProcessingRequired;
    }

    @Override
    public void setFormulaProcessingRequired(boolean formulaProcessingRequired) {
        this.formulaProcessingRequired = formulaProcessingRequired;
    }

    @Override
    public boolean isIgnoreSourceCellStyle() {
        return ignoreSourceCellStyle;
    }

    @Override
    public void setIgnoreSourceCellStyle(boolean ignoreSourceCellStyle) {
        this.ignoreSourceCellStyle = ignoreSourceCellStyle;
    }

    @Override
    public Map<String, String> getCellStyleMap() {
        return cellStyleMap;
    }

    @Override
    public void setCellStyleMap(Map<String, String> cellStyleMap) {
        this.cellStyleMap = cellStyleMap;
    }
}
