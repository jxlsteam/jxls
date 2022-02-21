package org.jxls.templatebasedtests;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.common.Context;
import org.jxls.entity.Employee;
import org.jxls.expression.EvaluationException;

public class Issue73Test {
    
    @Test(expected = EvaluationException.class)
    public void simple() {
        Context context = new Context();
        context.putVar("ff", "abc");
        
        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass(), "simple");
        tester.processTemplate(context);
    }

    @Test(expected = EvaluationException.class)
    public void each() {
        Context context = new Context();
        context.putVar("employees", Employee.generateSampleEmployeeData());

        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass(), "each");
        tester.processTemplate(context);
    }
}
