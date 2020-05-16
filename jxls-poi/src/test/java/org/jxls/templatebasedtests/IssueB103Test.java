package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;

import java.util.ArrayList;
import java.util.Arrays;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;

/**
 * Cell format not being correctly shifted when having JXLS command and empty list
 */
public class IssueB103Test {

    @Test
    public void test() {
        // Prepare
        Context context = new Context();
        context.putVar("nonEmptyList", new ArrayList<>(Arrays.asList("A", "B", "C", "D", "E", "F", "G", "H", "I", "J")));
        context.putVar("emptyList", new ArrayList<String>());
        
        // Test
        JxlsTester tester = JxlsTester.xls(getClass());
        tester.processTemplate(context);
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet(0);
            assertEquals(1, w.getCellValueAsDouble(3, 1), 0.01d);
            assertEquals("1: This is a very long text", w.getCellValueAsString(4, 1));
            assertEquals(2, w.getCellValueAsDouble(6, 1), 0.01d);
            assertEquals("2: This is a very long text", w.getCellValueAsString(7, 1));
            assertEquals("A", w.getCellValueAsString(8, 1));
            assertEquals("B", w.getCellValueAsString(9, 1)); // ...
            assertEquals("J", w.getCellValueAsString(17, 1));
            for (int row = 20; row <= 29; row++) {
                assertEquals("a large block afterwards", w.getCellValueAsString(row, 1));
            }
        }
    }
}
