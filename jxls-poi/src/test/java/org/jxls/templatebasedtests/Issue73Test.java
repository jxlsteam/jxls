package org.jxls.templatebasedtests;

import org.junit.Assert;
import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.JxlsTester.TransformerChecker;
import org.jxls.common.Context;
import org.jxls.common.JxlsException;
import org.jxls.common.PoiExceptionThrower;
import org.jxls.entity.Employee;
import org.jxls.expression.JexlExpressionEvaluator;
import org.jxls.transform.Transformer;

/**
 * Use PoiExceptionThrower to throw JXLS and JEXL exceptions instead of just logging them.
 */
public class Issue73Test {
    // Let's test a simple JEXL expression inside a field and a expression inside a jx:each command.
    // Test it with a syntax error and also with an unknown field name. So we have 4 testcases.
    // There are also 2 additional testcases.

    /**
     * simple field with syntax error
     */
    @Test(expected = JxlsException.class)
    public void simpleError() {
        // Prepare
        Context context = new Context();
        context.putVar("ff", "abc");
        
        // Test
        check(context, "simpleError");
    }

    /**
     * extra test: user needs to know what's wrong
     */
    @Test
    public void testExceptionMessage() {
        // Prepare
        Context context = new Context();
        context.putVar("ff", "abc");
        
        // Test
        try {
            check(context, "simpleError");
            
            // Verify
            Assert.fail("JxlsException with message \"...ff:0...\" expected!");
        } catch (JxlsException e) {
            Assert.assertTrue("\"ff:0\" should be mentioned", e.getMessage().contains("ff:0"));
            Assert.assertTrue("\"Tabelle1!C3\" should be mentioned", e.getMessage().contains("Tabelle1!C3"));
        }
    }

    /**
     * simple field which name is unknown
     */
    @Test(expected = JxlsException.class)
    public void simpleUnknown() {
        // Prepare
        Context context = new Context();
        // no "ff"
        
        // Test
        check(context, "simpleUnknown");
    }

    /**
     * jx:each with syntax error
     */
    @Test(expected = JxlsException.class)
    public void eachError() {
        Context context = new Context();
        context.putVar("employees", Employee.generateSampleEmployeeData());

        // Test
        check(context, "eachError");
    }

    /**
     * jx:each with unknown field
     */
    @Test(expected = JxlsException.class)
    public void eachUnknown() {
        Context context = new Context();
        context.putVar("employees", Employee.generateSampleEmployeeData());

        // Test
        check(context, "eachUnknown");
    }

    /**
     * Throw exception instead of logging a warning in EachCommand if expression is wrong.
     */
    @Test(expected = JxlsException.class)
    public void eachCollection() {
        Context context = new Context();
        context.putVar("employees", Employee.generateSampleEmployeeData());

        // Test
        check(context, "eachCollection");
    }

    private void check(Context context, String method) {
        JxlsTester tester = JxlsTester.xlsx(getClass(), method);
        tester.createTransformerAndProcessTemplate(context, new TransformerChecker() {
            @Override
            public Transformer checkTransformer(Transformer transformer) {
                // strict non-silent mode for getting all errors
                transformer.getTransformationConfig().setExpressionEvaluatorFactory(x -> new JexlExpressionEvaluator(false, true));
                
                // throw exceptions instead of just logging them
                transformer.setExceptionHandler(new PoiExceptionThrower());
                return transformer;
            }
        });
    }
}
