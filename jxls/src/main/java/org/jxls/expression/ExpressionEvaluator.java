package org.jxls.expression;

import java.util.Map;

/**
 * An interface to evaluate expressions
 * 
 * @author Leonid Vysochyn
 */
public interface ExpressionEvaluator {

    Object evaluate(String expression, Map<String,Object> context);
    
    Object evaluate(Map<String, Object> context);
    
    String getExpression();
}
