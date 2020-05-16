package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;
import org.jxls.entity.Employee;

/**
 * How can I get an index in jx:each?
 * 
 * varIndex new in jx:each
 */
public class IssueB173Test {

    @Test
    public void test() throws Exception {
        // Prepare
        Context context = new Context();
        context.putVar("employees", Employee.generateSampleEmployeeData());

        // Test
        JxlsTester tester = JxlsTester.xls(getClass());
        tester.processTemplate(context);
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet("Template");
            // check index in column E starts with row E
            for (int i = 0; i <= 4; i++) {
                assertEquals(i, w.getCellValueAsDouble(4 + i, 5), 0.1d);
            }
        }
    }
}
