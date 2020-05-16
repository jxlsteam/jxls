package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;

import java.util.Arrays;
import java.util.List;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;

/**
 * @see IssueB089Test
 * @see IssueB110Test
 */
public class ConditionalFormattingTest {
    private static final double EPSILON = 0.001;

    @Test
    public void shouldCopyConditionalFormatInEachCommandLoop() {
        // Prepare
        Context context = new Context();
        List<Integer> numbers = Arrays.asList(2, 1, 4, 3, 5);
        context.putVar("numbers", numbers);
        context.putVar("val1", 0);
        context.putVar("val2", 7);
        
        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);

        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet(0);
            for (int i = 0; i < numbers.size(); i++) {
                double val = w.getCellValueAsDouble(i + 2, 2);
                assertEquals(numbers.get(i).doubleValue(), val, EPSILON);
            }

            assertEquals(numbers.size() * 2 // for B2 list in template
                    + 2, // for A3:B3 in template (val1, val2)
                    w.getConditionalFormattingSize());

        }
    }
}
