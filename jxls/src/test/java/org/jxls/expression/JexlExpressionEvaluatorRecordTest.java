package org.jxls.expression;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

import java.util.HashMap;
import java.util.Map;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;

/**
 * Mirrors {@link JexlExpressionEvaluatorTest} using a Java Record instead of {@link Dummy}.
 * {@link JexlExpressionEvaluator} by default.
 *
 * @see JexlExpressionEvaluatorTest
 */
public class JexlExpressionEvaluatorRecordTest {

    /** Record equivalent of {@link Dummy}: {@code strValue} / {@code intValue} via accessor methods. */
    public record DummyRecord(String strValue, int intValue) {
    }

    @Before
    public void setUp() {
        JexlExpressionEvaluator.clear();
    }

    @After
    public void tearDown() {
        JexlExpressionEvaluator.clear();
    }

    @Test
    public void recordPropertyAccess() {
        String expression = "dummy.intValue";
        Map<String, Object> vars = new HashMap<>();
        vars.put("dummy", new DummyRecord("hello", 42));
        ExpressionEvaluator expressionEvaluator = new JexlExpressionEvaluator();
        Object result = expressionEvaluator.evaluate(expression, vars);
        assertNotNull(result);
        assertEquals(42, result);
    }

    @Test
    public void recordArithmeticWithProperty() {
        String expression = "2*x + dummy.intValue";
        Map<String, Object> vars = new HashMap<>();
        vars.put("x", Integer.valueOf(2));
        vars.put("dummy", new DummyRecord("hello", 3));                  // same arithmetic as original but with record
        ExpressionEvaluator expressionEvaluator = new JexlExpressionEvaluator();
        Object result = expressionEvaluator.evaluate(expression, vars);
        assertNotNull(result);
        assertEquals("7", result.toString());                            // 2*2 + 3 = 7, same as original simple2VarExpression
    }

    @Test
    public void recordStringProperty() {
        String expression = "dummy.strValue";
        Map<String, Object> vars = new HashMap<>();
        vars.put("dummy", new DummyRecord("hello", 42));
        ExpressionEvaluator expressionEvaluator = new JexlExpressionEvaluator();
        Object result = expressionEvaluator.evaluate(expression, vars);
        assertNotNull(result);
        assertEquals("hello", result);
    }
}
