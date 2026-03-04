package org.jxls.expression;

/**
 * A factory to create {@link ExpressionEvaluator} instance implementation based on JEXL
 */
public class ExpressionEvaluatorFactoryJexlImpl implements ExpressionEvaluatorFactory {
	private final boolean silent;
	private final boolean strict;
	private final JxlsJexlPermissions permissions;
    private JexlContextFactory jexlContextFactory;

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

	public ExpressionEvaluatorFactoryJexlImpl(boolean silent, boolean strict, JxlsJexlPermissions permissions, JexlContextFactory jexlContextFactory) {
		this.silent = silent;
		this.strict = strict;
		this.permissions = permissions;
		this.jexlContextFactory = jexlContextFactory;
	}

	@Override
    public ExpressionEvaluator createExpressionEvaluator(final String expression) {
		JexlExpressionEvaluator ee = expression == null ? new JexlExpressionEvaluator(silent, strict, permissions)
				: new JexlExpressionEvaluator(silent, strict, permissions, expression);
        if (jexlContextFactory != null) {
            ee.setJexlContextFactory(jexlContextFactory);
        }
        return ee;
    }

    public JexlContextFactory getJexlContextFactory() {
        return jexlContextFactory;
    }

    public void setJexlContextFactory(JexlContextFactory jexlContextFactory) {
        this.jexlContextFactory = jexlContextFactory;
    }
}
