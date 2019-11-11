package org.jxls.templatebasedtests;

import java.util.List;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.common.Context;
import org.jxls.entity.Employee;

public class Issue153Test {

    @Test
    public void test() throws Exception {
        Context context = new Context();
        List<Employee> employees = Employee.generateSampleEmployeeData();
        context.putVar("employees", employees);

        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);
        
        // TODO assertions
    }
}