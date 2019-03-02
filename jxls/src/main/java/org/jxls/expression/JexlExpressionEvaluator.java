package org.jxls.expression;

import org.apache.commons.jexl3.JexlBuilder;
import org.apache.commons.jexl3.JexlContext;
import org.apache.commons.jexl3.JexlEngine;
import org.apache.commons.jexl3.JexlExpression;
import org.apache.commons.jexl3.MapContext;

import java.util.HashMap;
import java.util.Map;

/**
 * @author Leonid Vysochyn
 */
public class JexlExpressionEvaluator implements ExpressionEvaluator{
    private JexlExpression jexlExpression;
    private JexlContext jexlContext;

    private ThreadLocal<JexlEngine> jexlThreadLocal;

    private ThreadLocal<Map<String, JexlExpression>> expressionMapThreadLocal = new ThreadLocal<Map<String, JexlExpression>>(){
        @Override
        protected Map<String, JexlExpression> initialValue() {
            return new HashMap<>();
        }
    };

    public JexlExpressionEvaluator() {
        this(true, false );
    }

    public JexlExpressionEvaluator(final boolean silent, final boolean strict) {
        jexlThreadLocal = new ThreadLocal<JexlEngine>() {
            @Override
            protected JexlEngine initialValue() {
                return new JexlBuilder().silent(silent).strict(strict).create();
            }
        };
    }

    public JexlExpressionEvaluator(String expression) {
        this();
        JexlEngine jexl = jexlThreadLocal.get();
        jexlExpression = jexl.createExpression( expression );
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
            JexlEngine jexl = jexlThreadLocal.get();
            Map<String,JexlExpression> expressionMap = expressionMapThreadLocal.get();
            JexlExpression jexlExpression =  expressionMap.get(expression);
            if( jexlExpression == null ){
                jexlExpression = jexl.createExpression( expression );
                expressionMap.put(expression, jexlExpression);
            }
            return jexlExpression.evaluate(jexlContext);
        } catch (Exception e) {
            throw new EvaluationException("An error occurred when evaluating expression " + expression, e);
        }
    }

    public Object evaluate(Map<String, Object> context){
        jexlContext = new MapContext(context);
        try {
            return jexlExpression.evaluate( jexlContext );
        } catch (Exception e) {
            throw new EvaluationException("An error occurred when evaluating expression " + jexlExpression.getSourceText(), e);
        }
    }

    public JexlExpression getJexlExpression() {
        return jexlExpression;
    }

    public void setJexlEngine(final JexlEngine jexlEngine){
        jexlThreadLocal = new ThreadLocal<JexlEngine>() {
            @Override
            protected JexlEngine initialValue() {
                return jexlEngine;
            }
        };
    }

    public JexlEngine getJexlEngine() {
        return jexlThreadLocal.get();
    }

	@Override
	public String getExpression() {
		return jexlExpression == null ? null : jexlExpression.getSourceText();
	}    
}
