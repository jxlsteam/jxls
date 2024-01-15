package org.jxls.templatebasedtests;

import java.util.ArrayList;
import java.util.Arrays;

import org.junit.Assert;
import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;
import org.jxls.common.ContextImpl;
import org.jxls.common.JxlsException;

/**
 * Checks whether EachCommand saves varIndex before the loop and restores it after the loop. (issues 93, 108)
 */
public class Issue93Test {

    @Test
    public void test() {
        // Prepare
        Context context = new ContextImpl() {
            @Override
            public Object getVar(String name) {
                if (name.substring(0, 1) == "m") { // Do an operation that could result in a NPE.
                    System.out.println("getVar: var starts with 'm'");
                }
                return super.getVar(name);
            }
        };
        context.putVar("list", new ArrayList<>(Arrays.asList("A", "B", "C")));
        context.putVar("list2", new ArrayList<>(Arrays.asList("D", "E")));
        context.putVar("i", -42d);
        
        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet(0);
            Assert.assertEquals("varIndex was not saved!", -42d, w.getCellValueAsDouble(7, 2), 0.1d);
            Assert.assertEquals("E", w.getCellValueAsString(10, 1));
        }
    }

    @Test
    public void wrongLastCellInTemplate() {
        Context context = new ContextImpl();
        context.putVar("list2", new ArrayList<>(Arrays.asList("D", "E")));
        JxlsTester tester = JxlsTester.xlsx(getClass(), "illegalArea");
        try {
            tester.processTemplate(context);
            Assert.fail("JxlsException expected!");
        } catch (JxlsException e) {
            Assert.assertTrue("Unexpected exception: " + e.getMessage(), e.getMessage().contains("Illegal area"));
        }
    }
}
