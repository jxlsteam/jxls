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
     * Evaluates if getExpression() is true.
     * @param context data access
     * @return expression result (true or false)
     * @throws JxlsException if return value is not a Boolean or null
     */
    default boolean isConditionTrue(Context context) {
        return isConditionTrue(getExpression(), context.toMap());
    }

    /**
     * Evaluates if getExpression() is true. Call this method only if you have no Context.
     * @param condition -
     * @param data -
     * @return expression result (true or false)
     * @throws JxlsException if return value is not a Boolean or null
     */
    default boolean isConditionTrue(String condition, Map<String, Object> data) {
        Object conditionResult = evaluate(condition, data);
        if (conditionResult instanceof Boolean b) {
            return Boolean.TRUE.equals(b);
        } else if (conditionResult == null) {
            throw new JxlsException("Result of condition \"" + condition + "\" is null");
        }
        throw new JxlsException("Result of condition \"" + condition + "\" is not a Boolean");
    }
}
