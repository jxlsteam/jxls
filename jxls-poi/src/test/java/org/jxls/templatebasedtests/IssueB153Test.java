package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;
import org.jxls.entity.Employee;

/**
 * Formula substitution does not work when referencing another sheet
 */
public class IssueB153Test {

    @Test
    public void test() throws Exception {
        // Prepare
        Context context = new Context();
        context.putVar("employees", Employee.generateSampleEmployeeData());

        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet("Template");
            assertEquals("EE", w.getCellValueAsString(4, 5));
            assertEquals("OO", w.getCellValueAsString(5, 5));
            assertEquals("NN", w.getCellValueAsString(6, 5));
            assertEquals("MM", w.getCellValueAsString(7, 5));
            assertEquals("JJ", w.getCellValueAsString(8, 5));
        }
    }
}