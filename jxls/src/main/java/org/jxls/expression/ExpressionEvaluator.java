package org.jxls.expression;

import java.util.Map;

import org.jxls.common.JxlsException;

/**
 * An interface to evaluate expressions
 * 
 * @author Leonid Vysochyn
 */
public interface ExpressionEvaluator {

    Object evaluate(String expression, Map<String, Object> context);
    
    Object evaluate(Map<String, Object> context);
    
    String getExpression();
    
    /**
     * Evaluates if the passed condition is true
     * @param condition expression
     * @param data -
     * @return expression result (true or false)
     * @throws JxlsException if return value is not a Boolean or null
     */
    default boolean isConditionTrue(String condition, Map<String, Object> data) {
        Object conditionResult = evaluate(condition, data);
        if (conditionResult instanceof Boolean b) {
            return Boolean.TRUE.equals(b);
        }
        throw new JxlsException("Result of condition \"" + condition + "\" is not a Boolean value or null");
    }
    
    /**
     * Evaluates if the passed condition is true
     * @param data -
     * @return expression result (true or false)
     * @throws JxlsException if return value is not a Boolean or null
     */
    default boolean isConditionTrue(Map<String, Object> data) {
        Object conditionResult = evaluate(data);
        if (conditionResult instanceof Boolean b) {
            return Boolean.TRUE.equals(b);
        }
        throw new JxlsException("Result of condition \"" + getExpression() + "\" is not a Boolean value or null");
    }
}
