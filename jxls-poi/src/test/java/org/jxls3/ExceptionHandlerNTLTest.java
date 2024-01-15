package org.jxls3;

import org.junit.After;
import org.junit.Before;
import org.jxls.expression.ExpressionEvaluatorFactory;
import org.jxls.expression.ExpressionEvaluatorNoThreadLocalFactoryJexlImpl;
import org.jxls.expression.JexlExpressionEvaluatorNoThreadLocal;
import org.jxls.expression.JxlsJexlPermissions;

/**
 * ExceptionHandler tests for ExpressionEvaluatorNoThreadLocalFactoryJexlImpl
 */
public class ExceptionHandlerNTLTest extends ExceptionHandlerTest {

	@Before
	@After
	@Override
	public void cleanup() {
		JexlExpressionEvaluatorNoThreadLocal.clear();
	}

	@Override
	protected ExpressionEvaluatorFactory getExpressionEvaluatorFactory(boolean strict) {
		return new ExpressionEvaluatorNoThreadLocalFactoryJexlImpl(strict);
	}

	@Override
	protected ExpressionEvaluatorFactory getExpressionEvaluatorFactory(boolean silent, boolean strict, JxlsJexlPermissions permissions) {
		return new ExpressionEvaluatorNoThreadLocalFactoryJexlImpl(silent, strict, permissions); 
	}
}
