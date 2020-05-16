package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;

import java.util.ArrayList;
import java.util.List;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;

/**
 * Issue 184
 * 
 * The template has a jx:if inside a jx:each. The sums of the columns before the jx:if-column are wrong.
 * 
 * Cause:
 * commit 5354beaf6c73bd252f98d272ac1e4dc8f870a13b,
 * issue #160 - add isForwardOnly() method to Transformer and processing static cells differently depending on this flag
 */
public class IssueB184Test {

    @Test
    public void test() {
        // Prepare
        List<Integer> data = new ArrayList<>();
        for (int j = 1; j <= 3; j++) {
            data.add(j);
        }
        Context context = new Context();
        context.putVar("data", data);
        
        // Test
        JxlsTester tester = JxlsTester.xls(getClass());
        tester.processTemplate(context);
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet("Bug");
            assertEquals(6d, w.getCellValueAsDouble(5, 3), 0.01d);
            assertEquals(6d, w.getCellValueAsDouble(5, 4), 0.01d);
            assertEquals(6d, w.getCellValueAsDouble(5, 5), 0.01d);
            assertEquals(6d, w.getCellValueAsDouble(5, 1), 0.01d);
            assertEquals(6d, w.getCellValueAsDouble(5, 2), 0.01d);
            assertEquals("SUM(C2:C4)", w.getFormulaString(5, 3));
       }
    }
}
