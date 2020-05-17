package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;

/**
 * Support for arrays in EachCommand
 */
public class IssueB198Test {

    @Test
    public void test() throws Exception {
        // Prepare
        String[] array = new String[] { "Arrays", "will", "work." };
        Context context = new Context();
        context.putVar("myArray", array);
        
        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet("Arrays");
            assertEquals("Arrays", w.getCellValueAsString(1, 1)); 
            assertEquals("will", w.getCellValueAsString(2, 1)); 
            assertEquals("work.", w.getCellValueAsString(3, 1)); 
        }
    }
}
