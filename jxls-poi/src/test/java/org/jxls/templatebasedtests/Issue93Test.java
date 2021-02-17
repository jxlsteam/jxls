package org.jxls.templatebasedtests;

import java.util.ArrayList;
import java.util.Arrays;

import org.junit.Assert;
import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;

/**
 * Checks whether EachCommand saves varIndex before the loop and restores it after the loop.
 */
public class Issue93Test {

    @Test
    public void test() {
        // Prepare
        Context context = new Context();
        context.putVar("list", new ArrayList<>(Arrays.asList("A", "B", "C")));
        context.putVar("i", -42d);
        
        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet(0);
            Assert.assertEquals("varIndex was not saved!", -42d, w.getCellValueAsDouble(7, 2), 0.1d);
        }
    }
}
