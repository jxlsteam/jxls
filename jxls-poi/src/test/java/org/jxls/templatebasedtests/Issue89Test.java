package org.jxls.templatebasedtests;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.common.Context;
import org.jxls.entity.Employee;

/**
 * Conditional formattings are not copied
 */
public class Issue89Test {

    @Test
    public void test() throws Exception {
        Context context = new Context();
        context.putVar("employees", Employee.generateSampleEmployeeData());
        
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);
        
        // TODO assertions
    }
}
