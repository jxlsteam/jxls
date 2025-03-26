package org.jxls.expression;

import java.util.HashMap;
import java.util.Map;

import org.apache.commons.jexl3.JexlBuilder;
import org.apache.commons.jexl3.JexlContext;
import org.apache.commons.jexl3.JexlEngine;
import org.apache.commons.jexl3.JexlExpression;
import org.apache.commons.jexl3.MapContext;
import org.apache.commons.jexl3.introspection.JexlPermissions;

/**
 * This is an implementation of {@link ExpressionEvaluator} without using {@link ThreadLocal}.
 * See issue B054 for more detail about the reason for it.
 * 
 * @author Leonid Vysochyn
 */
public class JexlExpressionEvaluatorNoThreadLocal implements ExpressionEvaluator {
    private static final Map<String, JexlExpression> expressionMap = new HashMap<>();
    private JexlExpression jexlExpression;
    private JexlContextFactory jexlContextFactory = context -> new MapContext(context);
    private JexlEngine jexl;

    public JexlExpressionEvaluatorNoThreadLocal() {
        this(true, false);
    }

    public JexlExpressionEvaluatorNoThreadLocal(boolean silent, boolean strict) {
        this(silent, strict, JxlsJexlPermissions.UNRESTRICTED);
    }

    public JexlExpressionEvaluatorNoThreadLocal(boolean silent, boolean strict, JxlsJexlPermissions permissions) {
        this(silent, strict, permissions.getJexlPermissions());
    }

    public JexlExpressionEvaluatorNoThreadLocal(boolean silent, boolean strict, JexlPermissions jexlPermissions) {
        this.jexl = new JexlBuilder().silent(silent).strict(strict).permissions(jexlPermissions).create();
    }

    public JexlExpressionEvaluatorNoThreadLocal(boolean silent, boolean strict, JxlsJexlPermissions permissions, String expression) {
        this(silent, strict, permissions.getJexlPermissions());
        jexlExpression = jexl.createExpression(expression);
    }

    public JexlExpressionEvaluatorNoThreadLocal(String expression) {
        this();
        jexlExpression = jexl.createExpression(expression);
    }

    @Override
    public Object evaluate(String expression, Map<String, Object> context) {
        JexlContext jexlContext = jexlContextFactory.create(context);
        try {
            JexlExpression aJexlExpression = expressionMap.get(expression);
            if (aJexlExpression == null) {
                aJexlExpression = jexl.createExpression(expression);
                expressionMap.put(expression, aJexlExpression);
            }
            return aJexlExpression.evaluate(jexlContext);
        } catch (Exception e) {
            throw new EvaluationException("An error occurred when evaluating expression " + expression, e);
            // JxlsLogger not needed here.
        }
    }

    @Override
    public Object evaluate(Map<String, Object> context) {
        JexlContext jexlContext = jexlContextFactory.create(context);
        try {
            return jexlExpression.evaluate(jexlContext);
        } catch (Exception e) {
            throw new EvaluationException("An error occurred when evaluating expression " + jexlExpression.getSourceText(), e);
            // JxlsLogger not needed here.
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
    
    /**
     * Clear expression cache
     */
    public static void clear() {
        expressionMap.clear();
    }

    public JexlContextFactory getJexlContextFactory() {
        return jexlContextFactory;
    }

    public void setJexlContextFactory(JexlContextFactory jexlContextFactory) {
        this.jexlContextFactory = jexlContextFactory;
    }
}
