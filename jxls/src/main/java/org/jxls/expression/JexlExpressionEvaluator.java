package org.jxls.expression;

import java.util.HashMap;
import java.util.Map;

import org.apache.commons.jexl3.JexlBuilder;
import org.apache.commons.jexl3.JexlContext;
import org.apache.commons.jexl3.JexlEngine;
import org.apache.commons.jexl3.JexlExpression;
import org.apache.commons.jexl3.MapContext;

/**
 * JEXL based implementation of {@link ExpressionEvaluator} interface
 * @author Leonid Vysochyn
 */
public class JexlExpressionEvaluator implements ExpressionEvaluator {
    private final boolean silent;
    private final boolean strict;
    private final JxlsJexlPermissions permissions;
    private JexlExpression jexlExpression;
    private JexlContext jexlContext;
    private static ThreadLocal<Map<String, JexlEngine>> jexlThreadLocal = new ThreadLocal<Map<String, JexlEngine>>() {
        @Override
        protected Map<String, JexlEngine> initialValue() {
            return new HashMap<String, JexlEngine>();
        }
    };
    private static final ThreadLocal<Map<String, JexlExpression>> expressionMapThreadLocal = new ThreadLocal<Map<String, JexlExpression>>() {
        @Override
        protected Map<String, JexlExpression> initialValue() {
            return new HashMap<>();
        }
    };

    public JexlExpressionEvaluator() {
        this(true, false);
    }

    public JexlExpressionEvaluator(final boolean silent, final boolean strict) {
        this(silent, strict, JxlsJexlPermissions.UNRESTRICTED);
    }
    
    public JexlExpressionEvaluator(boolean silent, boolean strict, JxlsJexlPermissions permissions) {
        this.silent = silent;
        this.strict = strict;
        this.permissions = permissions;
    }

    public JexlExpressionEvaluator(boolean silent, boolean strict, JxlsJexlPermissions permissions, String expression) {
        this(silent, strict, permissions);
        jexlExpression = getJexlEngine().createExpression(expression);
    }

    public JexlExpressionEvaluator(String expression) {
        this();
        jexlExpression = getJexlEngine().createExpression(expression);
    }

    public JexlExpressionEvaluator(Map<String, Object> context) {
        this();
        jexlContext = new MapContext(context);
    }

    public JexlExpressionEvaluator(JexlContext jexlContext) {
        this();
        this.jexlContext = jexlContext;
    }

    @Override
    public Object evaluate(String expression, Map<String, Object> context) {
        JexlContext jexlContext = new MapContext(context);
        try {
            Map<String, JexlExpression> expressionMap = expressionMapThreadLocal.get();
            JexlExpression jexlExpression = expressionMap.get(expression);
            if (jexlExpression == null) {
                jexlExpression = getJexlEngine().createExpression(expression);
                expressionMap.put(expression, jexlExpression);
            }
            return jexlExpression.evaluate(jexlContext);
        } catch (Exception e) {
            throw new EvaluationException("An error occurred when evaluating expression " + expression, e); // TODO JxlsLogger?
        }
    }

    @Override
    public Object evaluate(Map<String, Object> context) {
        jexlContext = new MapContext(context);
        try {
            return jexlExpression.evaluate(jexlContext);
        } catch (Exception e) {
            throw new EvaluationException("An error occurred when evaluating expression " + jexlExpression.getSourceText(), e); // TODO JxlsLogger?
        }
    }

    public JexlExpression getJexlExpression() {
        return jexlExpression;
    }

    public void setJexlEngine(JexlEngine jexlEngine) {
        String key = key();
        Map<String, JexlEngine> map = jexlThreadLocal.get();
        map.put(key, jexlEngine);
    }

    public JexlEngine getJexlEngine() {
        String key = key();
        Map<String, JexlEngine> map = jexlThreadLocal.get();
        JexlEngine ret = map.get(key);
        if (ret == null) {
            ret = new JexlBuilder().silent(silent).strict(strict).permissions(permissions.getJexlPermissions()).create();
            map.put(key, ret);
        }
        return ret;
    }

    private String key() {
        return silent + "-" + strict + "-" + permissions.hashCode();
    }

    @Override
    public String getExpression() {
        return jexlExpression == null ? null : jexlExpression.getSourceText();
    }

    /**
     * Clear expression cache for current thread
     */
    public static void clear() {
        expressionMapThreadLocal.get().clear();
        jexlThreadLocal.get().clear();
    }
}
