package org.jxls.templatebasedtests;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.common.Context;
import org.jxls.entity.Employee;

/**
 * How can I get an index in jx:each?
 * 
 * varIndex new in jx:each
 */
public class Issue173Test {

    @Test
    public void test() throws Exception {
        Context context = new Context();
        context.putVar("employees", Employee.generateSampleEmployeeData());

        JxlsTester tester = JxlsTester.xls(getClass());
        tester.processTemplate(context);
        
        // TODO Check testcase!
        // TODO assertions
    }
}
