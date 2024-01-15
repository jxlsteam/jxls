package org.jxls.expression;

/**
 * A factory to create {@link JexlExpressionEvaluatorNoThreadLocal} instance implementation based on JEXL
 */
public class ExpressionEvaluatorNoThreadLocalFactoryJexlImpl implements ExpressionEvaluatorFactory {
	private final boolean silent;
	private final boolean strict;
	private final JxlsJexlPermissions permissions;

	public ExpressionEvaluatorNoThreadLocalFactoryJexlImpl() {
		this(false);
	}

	public ExpressionEvaluatorNoThreadLocalFactoryJexlImpl(boolean strict) {
		this(!strict, strict, JxlsJexlPermissions.UNRESTRICTED);
	}
	
	public ExpressionEvaluatorNoThreadLocalFactoryJexlImpl(boolean silent, boolean strict, JxlsJexlPermissions permissions) {
		this.silent = silent;
		this.strict = strict;
		this.permissions = permissions;
	}

	@Override
    public ExpressionEvaluator createExpressionEvaluator(final String expression) {
        return expression == null ? new JexlExpressionEvaluatorNoThreadLocal(silent, strict, permissions)
        		                  : new JexlExpressionEvaluatorNoThreadLocal(silent, strict, permissions, expression);
    }
}
