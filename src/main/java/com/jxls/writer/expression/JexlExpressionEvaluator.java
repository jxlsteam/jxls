package com.jxls.writer.expression;




import org.apache.commons.jexl2.Expression;
import org.apache.commons.jexl2.JexlContext;
import org.apache.commons.jexl2.JexlEngine;
import org.apache.commons.jexl2.MapContext;

import java.util.Iterator;
import java.util.Map;

/**
 * Date: Nov 2, 2009
 *
 * @author Leonid Vysochyn
 */
public class JexlExpressionEvaluator implements ExpressionEvaluator{
    Map<String, Object> context;

    public JexlExpressionEvaluator(Map<String, Object> context) {
        this.context = context;
    }

    public Object evaluate(String expression) {
        try {
            JexlEngine jexl = new JexlEngine();
            Expression jexlExpression = jexl.createExpression( expression );
            JexlContext jexlContext = new MapContext();
            for(Map.Entry<String, Object> entry : context.entrySet()){
                jexlContext.set(entry.getKey(), entry.getValue());
            }
            return jexlExpression.evaluate(jexlContext);
        } catch (Exception e) {
            throw new EvaluationException("An error occurred when evaluating expression " + expression, e);
        }
    }
}
