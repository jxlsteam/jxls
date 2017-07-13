package org.jxls.expression;

public class ExpressionEvaluatorFactoryJexlImpl implements ExpressionEvaluatorFactory {

	@Override
	public ExpressionEvaluator createExpressionEvaluator(final String expression) {
		return expression == null ?  new JexlExpressionEvaluator() :  new JexlExpressionEvaluator(expression);
	}

}
