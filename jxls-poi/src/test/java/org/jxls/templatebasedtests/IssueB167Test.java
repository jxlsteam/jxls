package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;

import java.io.IOException;
import java.util.Date;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;
import org.jxls.entity.Employee;

/**
 * Excel templates context variable not found if used beyond column AZ
 */
public class IssueB167Test {

    @Test
    public void test() throws IOException {
        // Prepare
        Context context = new Context();
        context.putVar("var1", "Variable 1");
        context.putVar("var2", new Employee("Leo", new Date(), 1000, 0.10));

        // Test
        JxlsTester tester = JxlsTester.xls(getClass());
        tester.processTemplate(context);

        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet(0);
            int ba = 26 /* A to Z */ + 26 /* AA to AZ */ + 1 /* BA */;
            assertEquals("Leo", w.getCellValueAsString(1, ba)); // BA1
            assertEquals(1000d, w.getCellValueAsDouble(1, ba + 2), 0.01d); // BC1
            assertEquals("Variable 1", w.getCellValueAsString(1, 1)); // A1
        }
    }
}
