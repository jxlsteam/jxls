package org.jxls.expression;

import org.apache.commons.jexl2.Expression;
import org.apache.commons.jexl2.JexlContext;
import org.apache.commons.jexl2.JexlEngine;
import org.apache.commons.jexl2.MapContext;

import java.util.HashMap;
import java.util.Map;

/**
 * This is an implementation of {@link ExpressionEvaluator} without using {@link ThreadLocal}.
 * See issue#54 for more detail about the reason for it
 * @author Leonid Vysochyn
 */
public class JexlExpressionEvaluatorNoThreadLocal implements ExpressionEvaluator{
    private Expression jexlExpression;
    private JexlContext jexlContext;

    private static final JexlEngine jexl = new JexlEngine();

    private static final Map<String, Expression> expressionMap = new HashMap<>();

    public JexlExpressionEvaluatorNoThreadLocal() {
    }

    public JexlExpressionEvaluatorNoThreadLocal(String expression) {
        jexlExpression = jexl.createExpression( expression );
    }

    public JexlExpressionEvaluatorNoThreadLocal(Map<String, Object> context) {
        jexlContext = new MapContext(context);
    }

    public JexlExpressionEvaluatorNoThreadLocal(JexlContext jexlContext) {
        this.jexlContext = jexlContext;
    }

    @Override
    public Object evaluate(String expression, Map<String, Object> context) {
        JexlContext jexlContext = new MapContext(context);
        try {
            Expression jexlExpression =  expressionMap.get(expression);
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
            throw new EvaluationException("An error occurred when evaluating expression " + jexlExpression.getExpression(), e);
        }
    }

    public Expression getJexlExpression() {
        return jexlExpression;
    }

    public JexlEngine getJexlEngine() {
        return jexl;
    }
    
	@Override
	public String getExpression() {
		return jexlExpression == null ? null : jexlExpression.getExpression();
	}
}
