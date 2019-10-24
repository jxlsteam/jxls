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
    private JexlExpression jexlExpression;
    private JexlContext jexlContext;
    private static ThreadLocal<Map<String, JexlEngine>> jexlThreadLocal = new ThreadLocal<Map<String, JexlEngine>>() {
        @Override
        protected Map<String, JexlEngine> initialValue() {
            return new HashMap<String, JexlEngine>();
        }
    };
    private static ThreadLocal<Map<String, JexlExpression>> expressionMapThreadLocal = new ThreadLocal<Map<String, JexlExpression>>() {
        @Override
        protected Map<String, JexlExpression> initialValue() {
            return new HashMap<>();
        }
    };

    public JexlExpressionEvaluator() {
        this(true, false);
    }

    public JexlExpressionEvaluator(final boolean silent, final boolean strict) {
        this.silent = silent;
        this.strict = strict;
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
            ret = new JexlBuilder().silent(silent).strict(strict).create();
            map.put(key, ret);
        }
        return ret;
    }

    private String key() {
        return silent + "-" + strict;
    }

    @Override
    public String getExpression() {
        return jexlExpression == null ? null : jexlExpression.getSourceText();
    }
}
