package org.jxls.expression;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

import java.util.HashMap;
import java.util.Map;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;

/**
 * Mirrors {@link JexlExpressionEvaluatorNoThreadLocalTest} using a Java Record instead of {@link Dummy}.
 * is built into {@link JexlExpressionEvaluatorNoThreadLocal} by default.
 *
 * @see JexlExpressionEvaluatorNoThreadLocalTest
 */
public class JexlExpressionEvaluatorNoThreadLocalRecordTest {

    /** Same record as in {@link JexlExpressionEvaluatorRecordTest}. */
    public record DummyRecord(String strValue, int intValue) {
    }

    @Before
    public void setUp() {
        JexlExpressionEvaluatorNoThreadLocal.clear();
    }

    @After
    public void tearDown() {
        JexlExpressionEvaluatorNoThreadLocal.clear();
    }

    @Test
    public void recordPropertyAccess() {
        String expression = "dummy.intValue";
        Map<String, Object> vars = new HashMap<>();
        vars.put("dummy", new DummyRecord("hello", 42));
        ExpressionEvaluator expressionEvaluator = new JexlExpressionEvaluatorNoThreadLocal();
        Object result = expressionEvaluator.evaluate(expression, vars);
        assertNotNull(result);
        assertEquals(42, result);
    }

    @Test
    public void recordArithmeticWithProperty() {
        String expression = "2*x + dummy.intValue";
        Map<String, Object> vars = new HashMap<>();
        vars.put("x", Integer.valueOf(2));
        vars.put("dummy", new DummyRecord("hello", 3));
        ExpressionEvaluator expressionEvaluator = new JexlExpressionEvaluatorNoThreadLocal();
        Object result = expressionEvaluator.evaluate(expression, vars);
        assertNotNull(result);
        assertEquals("7", result.toString());
    }

    @Test
    public void recordStringProperty() {
        String expression = "dummy.strValue";
        Map<String, Object> vars = new HashMap<>();
        vars.put("dummy", new DummyRecord("hello", 42));
        ExpressionEvaluator expressionEvaluator = new JexlExpressionEvaluatorNoThreadLocal();
        Object result = expressionEvaluator.evaluate(expression, vars);
        assertNotNull(result);
        assertEquals("hello", result);
    }
}
