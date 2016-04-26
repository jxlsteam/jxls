package org.jxls.expression;




import org.apache.commons.jexl2.Expression;
import org.apache.commons.jexl2.JexlContext;
import org.apache.commons.jexl2.JexlEngine;
import org.apache.commons.jexl2.MapContext;

import java.util.HashMap;
import java.util.Map;

/**
 * Date: Nov 2, 2009
 *
 * @author Leonid Vysochyn
 */
public class JexlExpressionEvaluator implements ExpressionEvaluator{
    private Expression jexlExpression;
    private JexlContext jexlContext;

    private static final ThreadLocal<JexlEngine> jexlThreadLocal = new ThreadLocal<JexlEngine>(){
        @Override
        protected JexlEngine initialValue() {
            return new JexlEngine();
        }
    };

    private static final ThreadLocal<Map<String, Expression>> expressionMapThreadLocal = new ThreadLocal<Map<String, Expression>>(){
        @Override
        protected Map<String, Expression> initialValue() {
            return new HashMap<>();
        }
    };

    public JexlExpressionEvaluator() {
    }

    public JexlExpressionEvaluator(String expression) {
        JexlEngine jexl = jexlThreadLocal.get();
        jexlExpression = jexl.createExpression( expression );
    }

    public JexlExpressionEvaluator(Map<String, Object> context) {
        jexlContext = new MapContext(context);
    }

    public JexlExpressionEvaluator(JexlContext jexlContext) {
        this.jexlContext = jexlContext;
    }

    @Override
    public Object evaluate(String expression, Map<String, Object> context) {
        JexlContext jexlContext = new MapContext(context);
        try {
            JexlEngine jexl = jexlThreadLocal.get();
            Map<String,Expression> expressionMap = expressionMapThreadLocal.get();
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
        return jexlThreadLocal.get();
    }
}
