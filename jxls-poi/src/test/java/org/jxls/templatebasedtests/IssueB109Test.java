package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;

import java.io.IOException;
import java.util.ArrayList;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;

/**
 * Issue with circular formula
 */
public class IssueB109Test {

    @Test
    public void test() throws IOException {
        // Prepare
        Context context = new Context();
        context.putVar("emptyList", new ArrayList<>());

        // Test
        JxlsTester tester = JxlsTester.xls(getClass());
        tester.dontEvaluateFormulas().processTemplate(context);

        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet(1); // 2nd sheet
            assertEquals(0, w.getCellValueAsDouble(2, 2), 0.001d);
        }
    }
}
