package org.jxls.expression;

/**
 * A factory to create {@link ExpressionEvaluator} instance implementation based on JEXL
 */
public class ExpressionEvaluatorFactoryJexlImpl implements ExpressionEvaluatorFactory {
    @Override
    public ExpressionEvaluator createExpressionEvaluator(final String expression) {
        return expression == null ? new JexlExpressionEvaluator() : new JexlExpressionEvaluator(expression);
    }
}
