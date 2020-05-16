package org.jxls.templatebasedtests;

import java.io.IOException;
import java.util.Arrays;
import java.util.List;

import org.junit.Assert;
import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;
import org.jxls.entity.Employee;
import org.jxls.examples.GridCommandDemo;

/**
 * Allow Excel formulas to work with jx:grid
 */
public class IssueB090Test {

    @Test
    public void test() throws IOException {
        // Prepare
        Context context = new Context();
        context.putVar("headers", Arrays.asList("Name", "Birthday", "Payment"));
        List<Employee> employees = Employee.generateSampleEmployeeData();
        List<List<Object>> data = GridCommandDemo.createGridData(employees);
        context.putVar("data", data);

        // Test
        JxlsTester tester = JxlsTester.xls(getClass());
        tester.processTemplate(context);
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet(0);
            Assert.assertEquals(1, w.getCellValueAsDouble(2, 2), 0.1d);
        }
    }
}
