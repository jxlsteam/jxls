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

/**
 * For the Jxls main part, expressions are used by cells, jx:if, jx:each/select and jx:each/groupBy.
 * These tests check whether all expressions are received by the ExpressionEvaluator.
 * This also checks if the configured ExpressionEvaluatorFactory has been set.
 */
public class ExpressionEvaluatorFactoryTest {

    @Test
    public void ifAndSelect_standard() throws IOException {
        ifAndSelect(STREAMING_OFF);
    }

    @Test
    public void ifAndSelect_streaming() throws IOException {
        ifAndSelect(AUTO_DETECT);
    }
    
    private void ifAndSelect(JxlsStreaming streaming) {
        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass());
		JxlsPoiTemplateFillerBuilder builder = JxlsPoiTemplateFillerBuilder.newInstance().withStreaming(streaming)
				.withExpressionEvaluatorFactory(new MyEvaluatorFactory(
						"employees", "e.payment<2500"/*select*/, "e.payment<2000"/*jx:if*/, "e.name", "e.birthDate", "e.payment"));
        tester.test(data(), builder);
        
        // Verify
        Assert.assertEquals(18, ((MyEvaluatorFactory) builder.getExpressionEvaluatorFactory()).used);
    }

	private Map<String, Object> data() {
		Map<String,Object> data = new HashMap<>();
        data.put("employees", Employee.generateSampleEmployeeData());
		return data;
	}

    @Test
    public void groupBy_standard() throws IOException {
    	groupBy(STREAMING_OFF);
    }

    @Test
    public void groupBy_streaming() throws IOException {
    	groupBy(AUTO_DETECT);
    }
    
    private void groupBy(JxlsStreaming streaming) {
        // Test
        Jxls3Tester tester = new Jxls3Tester(GroupByTest.class, "GroupByTest_asc.xlsx"); // reuse template from other testcase
        JxlsPoiTemplateFillerBuilder builder = JxlsPoiTemplateFillerBuilder.newInstance().withStreaming(streaming)
        		.withExpressionEvaluatorFactory(new MyEvaluatorFactory(
        				"employees", "g.item.salaryGroup", "g.items", "e.name", "e.payment"));
        tester.test(data(), builder);
        
        // Verify
        Assert.assertEquals(15, ((MyEvaluatorFactory) builder.getExpressionEvaluatorFactory()).used);
    }

    static class MyEvaluatorFactory implements ExpressionEvaluatorFactory {
        private final ExpressionEvaluatorFactory parent = new ExpressionEvaluatorFactoryJexlImpl();
        private final Set<String> allowedExpressions = new HashSet<>();
        int used = 0;
        
        MyEvaluatorFactory(String ...pAllowedExpressions) {
        	for (String i : pAllowedExpressions) {
        		allowedExpressions.add(i);
        	}
        }

        @Override
        public ExpressionEvaluator createExpressionEvaluator(String topExpression) {
            ExpressionEvaluator parentEE = parent.createExpressionEvaluator(topExpression);
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
                public Object evaluate(Map<String, Object> context) {
                    if (allowedExpressions.contains(topExpression) ) {
                        used++;
                    } else {
                        Assert.fail("Unexpected expression: " + topExpression);
                    }
                    return parentEE.evaluate(context);
                }
                
                @Override
                public String getExpression() {
                    Assert.fail("unexpected call");
                    return null;
                }
            };
        }
    }
}
