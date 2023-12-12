package org.jxls3;

import static org.jxls.builder.JxlsStreaming.AUTO_DETECT;
import static org.jxls.builder.JxlsStreaming.STREAMING_OFF;

import java.io.IOException;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

import org.junit.Assert;
import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.builder.JxlsStreaming;
import org.jxls.entity.Employee;
import org.jxls.expression.ExpressionEvaluator;
import org.jxls.expression.ExpressionEvaluatorFactory;
import org.jxls.expression.ExpressionEvaluatorFactoryJexlImpl;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

// TODO work in progress
public class ExpressionEvaluatorFactoryTest {

    @Test
    public void standard() throws IOException {
        check(STREAMING_OFF);
    }

//TODO    @Test
    public void streaming() throws IOException {
        check(AUTO_DETECT);
    }
    
    private void check(JxlsStreaming streaming) {
        // Prepare
        Map<String,Object> data = new HashMap<>();
        data.put("employees", Employee.generateSampleEmployeeData());

        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass());
        JxlsPoiTemplateFillerBuilder builder = JxlsPoiTemplateFillerBuilder.newInstance().withStreaming(streaming).withExpressionEvaluatorFactory(new MyEvaluatorFactory());
        tester.test(data, builder);
        // TODO Test mit select
        
        // Verify
        Assert.assertEquals(21, ((MyEvaluatorFactory) builder.getExpressionEvaluatorFactory()).used);
    }
    
    // TODO Test mit select + group by
    
    static class MyEvaluatorFactory implements ExpressionEvaluatorFactory {
        private final ExpressionEvaluatorFactory parent = new ExpressionEvaluatorFactoryJexlImpl();
        private final Set<String> allowedExpressions = new HashSet<>();
        int used = 0;
        
        MyEvaluatorFactory() {
            allowedExpressions.add("employees");
            allowedExpressions.add("e.payment<2000");
            allowedExpressions.add("e.name");
            allowedExpressions.add("e.birthDate");
            allowedExpressions.add("e.payment");
        }

        @Override
        public ExpressionEvaluator createExpressionEvaluator(String expression) {
            ExpressionEvaluator parentEE = parent.createExpressionEvaluator(expression);
            return new ExpressionEvaluator() {
                @Override
                public Object evaluate(String expression, Map<String, Object> context) {
                    if (allowedExpressions.contains(expression) ) {
                        used++;
                    } else {
                        Assert.fail("Unexpected expression: " + expression);
                    }
                    return parentEE.evaluate(expression, context);
                }
                
                @Override
                public String getExpression() {
                    Assert.fail("unexpected call");
                    return null;
                }
                
                @Override
                public Object evaluate(Map<String, Object> context) {
                    Assert.fail("unexpected call");
                    return null;
                }
            };
        }
    }
}
