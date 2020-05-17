package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;

import java.io.IOException;
import java.util.Arrays;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;

/**
 * Issue with 'big double values' (like 1.3E22) being parsed as cell references
 * (like E22)
 */
public class IssueB105Test {

    @Test
    public void test() throws IOException {
        // Prepare
        Context context = new Context();
        context.putVar("vars", Arrays.asList(new Values(), new Values(), new Values()));

        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);

        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet(1); // 2nd sheet
            for (int row = 22; row <= 24; row++) {
                assertEquals(1.3E22, w.getCellValueAsDouble(row, 2), 1d);  // B
                assertEquals(2.6E22, w.getCellValueAsDouble(row, 3), 1d);  // C
                assertEquals(2.6E22, w.getCellValueAsDouble(row, 4), 1d);  // D
                assertEquals(1.69E44, w.getCellValueAsDouble(row, 5), 1d); // E
                assertEquals(0, w.getCellValueAsDouble(row, 6), 1d);       // F
                assertEquals(1.56E22, w.getCellValueAsDouble(row, 7), 1d); // G
                assertEquals(1.56E22, w.getCellValueAsDouble(row, 8), 1d); // H
                assertEquals(1.3E22, w.getCellValueAsDouble(row, 9), 1d);  // I
                assertEquals(1.3E22, w.getCellValueAsDouble(row, 10), 1d); // J
            }
        }
    }

    public static class Values {
        public double smallValue = 1.2;
        public double bigValue = 1.3E22;
    }
}