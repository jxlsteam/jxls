package org.jxls.templatebasedtests;

import java.util.List;

import org.junit.Assert;
import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;
import org.jxls.entity.Employee;

/**
 * Conditional formattings are not copied
 * 
 * <p>note: Issue B089 means issue 89 from BitBucket. If there's no B prefix it's a Github issue number.</p>
 * 
 * @see ConditionalFormattingTest
 * @see IssueB110Test
 */
public class IssueB089Test {

    @Test
    public void test() throws Exception {
        // Prepare
        Context context = new Context();
        List<Employee> employees = Employee.generateSampleEmployeeData();
        context.putVar("employees", employees);
        
        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet(0);
            Assert.assertEquals("Conditional formattings have not been copied!", employees.size(), w.getConditionalFormattingSize());
        }
    }
}
