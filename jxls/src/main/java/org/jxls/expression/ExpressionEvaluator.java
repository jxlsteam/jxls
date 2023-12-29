package org.jxls.expression;

import java.util.Map;

import org.jxls.common.Context;
import org.jxls.common.JxlsException;

/**
 * An interface to evaluate expressions
 * 
 * @author Leonid Vysochyn
 */
public interface ExpressionEvaluator {

    Object evaluate(String expression, Map<String, Object> data);
    
    Object evaluate(Map<String, Object> data);
    
    String getExpression();
    
    /**
     * Evaluates if getExpression() is true
     * @param context data access
     * @return expression result (true or false)
     * @throws JxlsException if return value is not a Boolean or null
     */
    default boolean isConditionTrue(Context context) {
        Object conditionResult = evaluate(context.toMap());
        if (conditionResult instanceof Boolean b) {
            return Boolean.TRUE.equals(b);
        } else if (conditionResult == null) {
            throw new JxlsException("Result of condition \"" + getExpression() + "\" is null");
        }
        throw new JxlsException("Result of condition \"" + getExpression() + "\" is not a Boolean");
    }
}
