package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;

import java.io.IOException;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;

/**
 * Tests orderBy attribute of jx:each command (Issue B193)
 */
public class OrderByTest {

    @Test
    public void test() throws IOException {
        // Prepare
        Context context = new Context();
        context.putVar("employees", IssueB133Test.createEmployees());
        
        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet("orderBy");
            assertEquals("03-1", w.getCellValueAsString(2, 1)); 
            assertEquals("Markus", w.getCellValueAsString(2, 2)); 

            assertEquals("03", w.getCellValueAsString(3, 1)); 
            assertEquals("Thomas", w.getCellValueAsString(3, 2)); 
            
            assertEquals("01", w.getCellValueAsString(4, 1)); 
            assertEquals("Herbert", w.getCellValueAsString(4, 2)); 
            
            assertEquals("01", w.getCellValueAsString(5, 1)); 
            assertEquals("Sven", w.getCellValueAsString(5, 2)); 
        }
    }
}
