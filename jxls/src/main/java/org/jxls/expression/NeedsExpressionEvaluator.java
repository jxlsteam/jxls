package org.jxls.expression;

public interface NeedsExpressionEvaluator {

    void setExpressionEvaluator(ExpressionEvaluator expressionEvaluator);
    
    // TODO feature can be removed because Context holds ExpressionEvaluator
}
