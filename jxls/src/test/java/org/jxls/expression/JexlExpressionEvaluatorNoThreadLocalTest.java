package org.jxls.expression;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;

import java.util.HashMap;
import java.util.Map;

import org.junit.Test;

/**
 * @author Leonid Vysochyn
 */
public class JexlExpressionEvaluatorNoThreadLocalTest {

    @Test
    public void simple2VarExpression() {
        String expression = "2 * x + y";
        Map<String, Object> vars = new HashMap<>();
        vars.put("x", Integer.valueOf(2));
        vars.put("y", Integer.valueOf(3));
        ExpressionEvaluator expressionEvaluator = new JexlExpressionEvaluatorNoThreadLocal();
        Object result = expressionEvaluator.evaluate( expression, vars );
        assertNotNull( result );
        assertEquals( "Simple 2-var expression evaluation result is wrong", "7", result.toString());
    }

    @Test
    public void shouldThrowEvaluationExceptionWhenError() {
        String expression = "2 * x + y )";
        Map<String, Object> vars = new HashMap<>();
        vars.put("x", Integer.valueOf(2));
        vars.put("y", Integer.valueOf(3));
        ExpressionEvaluator expressionEvaluator = new JexlExpressionEvaluatorNoThreadLocal();
        try {
            expressionEvaluator.evaluate(expression, vars);
            fail("EvaluationException expected");
        } catch (EvaluationException expected) {
            assertTrue(expected.getMessage().contains("error"));
            assertTrue(expected.getMessage().contains(expression));
        }
    }

    @Test
    public void evaluateWhenVarIsNull() {
        String expression = "2*x + dummy.intValue";
        Map<String, Object> vars = new HashMap<>();
        vars.put("x", Integer.valueOf(2));
        vars.put("dummy", null);
        ExpressionEvaluator expressionEvaluator = new JexlExpressionEvaluatorNoThreadLocal();
        Object result = expressionEvaluator.evaluate( expression , vars);
        assertEquals("Incorrect evaluation when a var is null", "4", result.toString());
    }
    
    @Test 
    public void evaluateWhenExpressionVarIsUndefined() {
        String expression = "dummy.intValue";
        Map<String, Object> vars = new HashMap<>();
        ExpressionEvaluator expressionEvaluator = new JexlExpressionEvaluatorNoThreadLocal();
        Object result = expressionEvaluator.evaluate( expression, vars );
        assertNull(result);
    }
}
