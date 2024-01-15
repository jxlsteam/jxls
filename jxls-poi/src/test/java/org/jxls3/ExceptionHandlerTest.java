package org.jxls3;

import static org.junit.Assert.assertTrue;

import java.util.HashMap;
import java.util.Map;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.common.JxlsException;
import org.jxls.entity.Employee;
import org.jxls.expression.ExpressionEvaluatorFactory;
import org.jxls.expression.ExpressionEvaluatorFactoryJexlImpl;
import org.jxls.expression.JexlExpressionEvaluator;
import org.jxls.expression.JxlsJexlPermissions;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

/**
 * ExceptionHandler tests for ExpressionEvaluatorFactoryJexlImpl
 */
public class ExceptionHandlerTest {

	@Before
	@After
	public void cleanup() {
		JexlExpressionEvaluator.clear();
	}
	
	@Test
	public void strictModeWithException() {
		Jxls3Tester tester = Jxls3Tester.xlsx(EachTest.class);
		JxlsPoiTemplateFillerBuilder builder = JxlsPoiTemplateFillerBuilder.newInstance().withExceptionThrower()
				.withExpressionEvaluatorFactory(getExpressionEvaluatorFactory(true));
		try {
			tester.test(new HashMap<>()/*error: we don't set employees*/, builder);
			Assert.fail("JxlsException expected");
		} catch (JxlsException expected) {
			assertTrue(expected.getMessage().contains("Failed to evaluate collection expression \"employees\""));
		}
	}

	@Test
	public void silentModeWithoutException() {
		Jxls3Tester tester = Jxls3Tester.xlsx(EachTest.class);
		JxlsPoiTemplateFillerBuilder builder = JxlsPoiTemplateFillerBuilder.newInstance().withExceptionThrower()
				.withExpressionEvaluatorFactory(getExpressionEvaluatorFactory(false));
		tester.test(new HashMap<>()/*error: we don't set employees*/, builder);
	}

	@Test
	public void noPermissions() {
		// Prepare
        Map<String,Object> data = new HashMap<>();
        data.put("employees", Employee.generateSampleEmployeeData());

        // Test
		Jxls3Tester tester = Jxls3Tester.xlsx(EachTest.class);
		JxlsPoiTemplateFillerBuilder builder = JxlsPoiTemplateFillerBuilder.newInstance().withExceptionThrower()
				.withExpressionEvaluatorFactory(getExpressionEvaluatorFactory(false, true, JxlsJexlPermissions.RESTRICTED));
		try {
			tester.test(data, builder);
			Assert.fail("JxlsException expected");
		} catch (JxlsException expected) {
			// Unfortunately, we don't get an error message saying that it's due to the permissions.
			assertTrue(expected.getMessage().contains("${e.name}"));
		}
	}
	
	@Test
	public void permissionsGranted() {
		// Prepare
        Map<String,Object> data = new HashMap<>();
        data.put("employees", Employee.generateSampleEmployeeData());

        // Test
		Jxls3Tester tester = Jxls3Tester.xlsx(EachTest.class);
		JxlsPoiTemplateFillerBuilder builder = JxlsPoiTemplateFillerBuilder.newInstance().withExceptionThrower()
				.withExpressionEvaluatorFactory(getExpressionEvaluatorFactory(false, true, new JxlsJexlPermissions(Employee.class.getName())));
		tester.test(data, builder);
	}
	
	protected ExpressionEvaluatorFactory getExpressionEvaluatorFactory(boolean strict) {
		return new ExpressionEvaluatorFactoryJexlImpl(strict);
	}

	protected ExpressionEvaluatorFactory getExpressionEvaluatorFactory(boolean silent, boolean strict, JxlsJexlPermissions permissions) {
		return new ExpressionEvaluatorFactoryJexlImpl(silent, strict, permissions);
	}
}
