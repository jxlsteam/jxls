package org.jxls.expression;

/**
 * A factory interface for creating {@link ExpressionEvaluator} instances
 */
public interface ExpressionEvaluatorFactory {
	ExpressionEvaluator createExpressionEvaluator(String expression);
}
