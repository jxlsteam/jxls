package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;
import org.jxls.entity.Employee;

/**
 * Bug in parameterized Excel Formulas
 * 
 * Template is quite similar to ParameterizedFormulasDemo. It contains the LEFT() Excel function in template cell A4.
 */
public class IssueB162Test {
    
    @Test
    public void test() {
        // Prepare
        Context context = new Context();
        context.putVar("employees", Employee.generateSampleEmployeeData());
        context.putVar("bonus", 0.1);

        // Test
        JxlsTester tester = JxlsTester.xls(getClass());
        tester.processTemplate(context);
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet(0);
            assertEquals("El", w.getCellValueAsString(4, 1));
            assertEquals("Jo", w.getCellValueAsString(8, 1));
        }
    }
}
