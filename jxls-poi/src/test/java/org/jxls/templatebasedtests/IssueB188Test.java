package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;

import java.io.IOException;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;

/**
 * A test case for Issue 188 Referencing other sheet in JXLS-processed cell formula replaces formula with "=0"
 */
public class IssueB188Test {

    @Test
    public void testCrossSheetFormulas() throws IOException {
        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(new Context());

        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet("Sheet1");
            assertEquals("Sheet2!A1", w.getFormulaString(1, 1));
            assertEquals("Static data", w.getCellValueAsString(1, 1));
        }
    }
}
