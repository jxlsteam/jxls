package org.jxls.expression;

/**
 * A factory to create {@link ExpressionEvaluator} instance implementation based on JEXL
 */
public class ExpressionEvaluatorFactoryJexlImpl implements ExpressionEvaluatorFactory {
	private final boolean silent;
	private final boolean strict;
	private final JxlsJexlPermissions permissions;

	public ExpressionEvaluatorFactoryJexlImpl() {
		this(false);
	}

	public ExpressionEvaluatorFactoryJexlImpl(boolean strict) {
		this(!strict, strict, JxlsJexlPermissions.UNRESTRICTED);
	}
	
	public ExpressionEvaluatorFactoryJexlImpl(boolean silent, boolean strict, JxlsJexlPermissions permissions) {
		this.silent = silent;
		this.strict = strict;
		this.permissions = permissions;
	}

	@Override
    public ExpressionEvaluator createExpressionEvaluator(final String expression) {
        return expression == null ? new JexlExpressionEvaluator(silent, strict, permissions)
        		                  : new JexlExpressionEvaluator(silent, strict, permissions, expression);
    }
}
