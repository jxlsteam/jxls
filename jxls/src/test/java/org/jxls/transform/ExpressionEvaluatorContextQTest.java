package org.jxls.transform;

import java.util.HashMap;
import java.util.Map;

import org.junit.Assert;
import org.junit.Test;
import org.jxls.builder.JxlsTemplateFillerBuilder;
import org.jxls.common.EvaluationResult;
import org.jxls.expression.ExpressionEvaluatorFactoryJexlImpl;

public class ExpressionEvaluatorContextQTest {
    // You get full test coverage by replacing ExpressionEvaluatorContext with ExpressionEvaluatorContextQ and running all test cases.
    
    @Test
    public void addition()  {
        Map<String, Object> data = new HashMap<>();
        data.put("fourty", Integer.valueOf(40));
        String expr = "${1+fourty}-30";
        ExpressionEvaluatorContext e = new ExpressionEvaluatorContextQ(new ExpressionEvaluatorFactoryJexlImpl(), null, null);
        
        EvaluationResult r = e.evaluateRawExpression(expr, data);
        
        Assert.assertEquals("41-30", r.getResult());
    }

    @Test
    public void callFunction()  {
        Map<String, Object> data = new HashMap<>();
        data.put("fo", new FuncObject());
        String expr = "${1+fo.func(\"fourty\")}-30";
        ExpressionEvaluatorContext e = new ExpressionEvaluatorContextQ(new ExpressionEvaluatorFactoryJexlImpl(), null, null);
        
        EvaluationResult r = e.evaluateRawExpression(expr, data);
        
        Assert.assertEquals("41-30", r.getResult());
    }

    @Test
    public void nested()  {
        Map<String, Object> data = new HashMap<>();
        data.put("fo", new FuncObject());
        String expr = "${1+fo.func2(\"${fourty}\")}-30";
        ExpressionEvaluatorContext e = new ExpressionEvaluatorContextQ(new ExpressionEvaluatorFactoryJexlImpl(), null, null);
        
        EvaluationResult r = e.evaluateRawExpression(expr, data);
        
        Assert.assertEquals("41-30", r.getResult());
    }

    @Test
    public void nested_backslashQuote()  {
        Map<String, Object> data = new HashMap<>();
        data.put("fo", new FuncObject());
        String expr = "${fo.func2(\"${\\\"fourty\\\"}\")}";
        ExpressionEvaluatorContext e = new ExpressionEvaluatorContextQ(new ExpressionEvaluatorFactoryJexlImpl(), null, null);
        
        EvaluationResult r = e.evaluateRawExpression(expr, data);
        
        Assert.assertEquals(80, r.getResult());
    }

    @Test
    public void two()  {
        Map<String, Object> data = new HashMap<>();
        data.put("fourty", "40");
        String expr = "${fourty}${fourty}";
        ExpressionEvaluatorContext e = new ExpressionEvaluatorContextQ(new ExpressionEvaluatorFactoryJexlImpl(), null, null);
        
        EvaluationResult r = e.evaluateRawExpression(expr, data);
        
        Assert.assertEquals("4040", r.getResult());
    }
    
    /**
     * Does the builder return the given ExpressionEvaluatorContextFactory instance?
     */
    @Test
    public void builder() {
        ExpressionEvaluatorContextFactory fac = (f, b, e) -> new ExpressionEvaluatorContextQ(f, b, e);
        
        JxlsTemplateFillerBuilder<?> builder = JxlsTemplateFillerBuilder.newInstance().withExpressionEvaluatorContextFactory(fac);
        ExpressionEvaluatorContextFactory factory = builder.getOptions().getExpressionEvaluatorContextFactory();
        
        Assert.assertEquals(ExpressionEvaluatorContextQ.class.getSimpleName(),
                factory.build(new ExpressionEvaluatorFactoryJexlImpl(), null, null).getClass().getSimpleName());
    }

    public static class FuncObject {
        
        public Integer func(String name) {
            return "fourty".equals(name) ? Integer.valueOf(40) : Integer.valueOf(0);
        }

        public Integer func2(String name) {
            if ("${\"fourty\"}".equals(name)) {
                return Integer.valueOf(80);
            } else if ("${fourty}".equals(name)) {
                return Integer.valueOf(40);
            }
            return Integer.valueOf(0);
        }
    }
}
