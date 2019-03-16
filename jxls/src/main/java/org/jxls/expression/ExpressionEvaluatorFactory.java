package org.jxls.expression;

public interface ExpressionEvaluatorFactory {

	ExpressionEvaluator createExpressionEvaluator(String expression);
}
