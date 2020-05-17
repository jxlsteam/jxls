package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;

/**
 * Wrong average on 2nd sheet
 */
public class IssueB166Test {

    @Test
    public void test() {
    	// Prepare: define result set
        List<Map<String, Object>> rs = new ArrayList<>();
        for (int i = 0; i < 5; i++) {
            Map<String, Object> map = new HashMap<String, Object>();
            map.put("count", i);
            rs.add(map);
        }
        final Context context = new Context();
        context.putVar("rs0", rs);

        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.dontEvaluateFormulas().processTemplate(context);
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            verifyTab(w, "Tab1");
            verifyTab(w, "Tab2");
        }
    }

    private void verifyTab(TestWorkbook w, String tabName) {
        w.selectSheet(tabName);
        assertEquals(1, w.getCellValueAsDouble(3, 1), 1e-3);
        assertEquals(4, w.getCellValueAsDouble(6, 1), 1e-3);
        assertEquals("AVERAGEA(A2:A6)", w.getFormulaString(7, 2));
        assertEquals("SUM(A2:A6)", w.getFormulaString(8, 2));
    }
}
