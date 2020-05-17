package org.jxls.expression;

import java.util.HashMap;
import java.util.Map;

import org.apache.commons.jexl3.JexlBuilder;
import org.apache.commons.jexl3.JexlContext;
import org.apache.commons.jexl3.JexlEngine;
import org.apache.commons.jexl3.JexlExpression;
import org.apache.commons.jexl3.MapContext;

/**
 * This is an implementation of {@link ExpressionEvaluator} without using {@link ThreadLocal}.
 * See issue B054 for more detail about the reason for it.
 * 
 * @author Leonid Vysochyn
 */
public class JexlExpressionEvaluatorNoThreadLocal implements ExpressionEvaluator {
    private static final Map<String, JexlExpression> expressionMap = new HashMap<>();
    private JexlExpression jexlExpression;
    private JexlContext jexlContext;
    private JexlEngine jexl;

    public JexlExpressionEvaluatorNoThreadLocal() {
        this(true, false);
    }

    public JexlExpressionEvaluatorNoThreadLocal(boolean silent, boolean strict) {
        this.jexl = new JexlBuilder().silent(silent).strict(strict).create();
    }

    public JexlExpressionEvaluatorNoThreadLocal(String expression) {
        this();
        jexlExpression = jexl.createExpression(expression);
    }

    public JexlExpressionEvaluatorNoThreadLocal(Map<String, Object> context) {
        this();
        jexlContext = new MapContext(context);
    }

    public JexlExpressionEvaluatorNoThreadLocal(JexlContext jexlContext) {
        this();
        this.jexlContext = jexlContext;
    }

    @Override
    public Object evaluate(String expression, Map<String, Object> context) {
        JexlContext jexlContext = new MapContext(context);
        try {
            JexlExpression jexlExpression = expressionMap.get(expression);
            if (jexlExpression == null) {
                jexlExpression = jexl.createExpression(expression);
                expressionMap.put(expression, jexlExpression);
            }
            return jexlExpression.evaluate(jexlContext);
        } catch (Exception e) {
            throw new EvaluationException("An error occurred when evaluating expression " + expression, e);
        }
    }

    @Override
    public Object evaluate(Map<String, Object> context) {
        jexlContext = new MapContext(context);
        try {
            return jexlExpression.evaluate(jexlContext);
        } catch (Exception e) {
            throw new EvaluationException("An error occurred when evaluating expression " + jexlExpression.getSourceText(), e);
        }
    }

    public JexlExpression getJexlExpression() {
        return jexlExpression;
    }

    public void setJexlEngine(final JexlEngine jexlEngine) {
        this.jexl = jexlEngine;
    }

    public JexlEngine getJexlEngine() {
        return jexl;
    }

    @Override
    public String getExpression() {
        return jexlExpression == null ? null : jexlExpression.getSourceText();
    }
}
