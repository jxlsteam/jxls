package org.jxls.expression;

import java.util.Map;

import org.jxls.common.JxlsException;

/**
 * An interface to evaluate expressions
 * 
 * @author Leonid Vysochyn
 */
public interface ExpressionEvaluator {

    Object evaluate(String expression, Map<String,Object> context);
    
    Object evaluate(Map<String, Object> context);
    
    String getExpression();
    
    /**
     * Evaluates if the passed condition is true
     * @param condition expression
     * @param data -
     * @return True, False or null
     * @throws JxlsException if return value is not a Boolean
     */
    default Boolean isConditionTrue(String condition, Map<String, Object> data) {
        Object conditionResult = evaluate(condition, data);
        if (conditionResult instanceof Boolean b) {
            return b;
        }
        throw new JxlsException("Result of condition \"" + condition + "\" is not a Boolean value");
    }
    
    /**
     * Evaluates if the passed condition is true
     * @param data -
     * @return true if the condition is evaluated to true or false otherwise
     */
    default Boolean isConditionTrue(Map<String, Object> data) {
        Object conditionResult = evaluate(data);
        if (conditionResult instanceof Boolean b) {
            return b;
        }
        throw new JxlsException("Result of condition \"" + getExpression() + "\" is not a Boolean value");
    }


}
