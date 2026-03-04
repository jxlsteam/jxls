package org.jxls3;

import java.io.FileNotFoundException;
import java.util.HashMap;
import java.util.Map;

import org.apache.commons.jexl3.JexlContext;
import org.apache.commons.jexl3.MapContext;
import org.junit.Assert;
import org.junit.Test;
import org.jxls.expression.ExpressionEvaluatorFactoryJexlImpl;
import org.jxls.expression.JexlContextFactory;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

/**
 * This testcase demonstrates how to call a top-level function (i.e., without an object) in a JEXL expression.
 * The function must be programmed in the class implementing JexlContext. By default, Jxls uses MapContext.
 * To replace this class, you must inject a JexlContextFactory. (issue 355)
 */
public class JexlContextFactoryTest {

    @Test
    public void test() throws FileNotFoundException {
    	var eeFactory = new ExpressionEvaluatorFactoryJexlImpl();
    	eeFactory.setJexlContextFactory(new MyJexlContextFactory());
		var builder = JxlsPoiTemplateFillerBuilder.newInstance().withExpressionEvaluatorFactory(eeFactory);
		var data = new HashMap<String, Object>();
		
		var result = (Double) builder.getExpressionEvaluatorFactory().createExpressionEvaluator("2+n111()").evaluate(data);

		Assert.assertEquals(113d, result.doubleValue(), 0.5d);
    }
    
    public static class MyJexlContextFactory implements JexlContextFactory {

		@Override
		public JexlContext create(Map<String, Object> context) {
			return new MyJexlContext(context);
		}
    }
    
    public static class MyJexlContext extends MapContext {

		public MyJexlContext(Map<String, Object> vars) {
			super(vars);
		}
		
		// JEXL function at root level
		public double n111() {
			return 111d;
		}
    }
}
