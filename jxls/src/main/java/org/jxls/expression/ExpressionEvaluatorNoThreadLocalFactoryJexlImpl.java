package org.jxls.expression;

/**
 * A factory to create {@link JexlExpressionEvaluatorNoThreadLocal} instance implementation based on JEXL
 */
public class ExpressionEvaluatorNoThreadLocalFactoryJexlImpl implements ExpressionEvaluatorFactory {
	private final boolean silent;
	private final boolean strict;
	private final JxlsJexlPermissions permissions;
    private JexlContextFactory jexlContextFactory;

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

	public ExpressionEvaluatorNoThreadLocalFactoryJexlImpl(boolean silent, boolean strict, JxlsJexlPermissions permissions, JexlContextFactory jexlContextFactory) {
		this.silent = silent;
		this.strict = strict;
		this.permissions = permissions;
		this.jexlContextFactory = jexlContextFactory;
	}

	@Override
    public ExpressionEvaluator createExpressionEvaluator(final String expression) {
		JexlExpressionEvaluatorNoThreadLocal ee = expression == null ? new JexlExpressionEvaluatorNoThreadLocal(silent, strict, permissions)
				: new JexlExpressionEvaluatorNoThreadLocal(silent, strict, permissions, expression);
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
