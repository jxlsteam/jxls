package org.jxls.command;

import static org.junit.Assert.assertEquals;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.junit.Test;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

/**
 * Tests orderBy attribute of jx:each command (Issue 193)
 */
public class OrderByTest {

    @Test
    public void test() throws IOException {
        // Prepare
        Context context = new Context();
        context.putVar("employees", Issue133Test.createEmployees());
        
        // Test
/*TODO  JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);
*/
        File outputFile = new File("target/OrderByTest_output.xlsx");
        try (InputStream in = getClass().getResourceAsStream("OrderByTest.xlsx")) {
            try (FileOutputStream out = new FileOutputStream(outputFile)) {
                JxlsHelper.getInstance().processTemplate(in, out, context);
            }
        }
        
        // Verify
//TODO  try (TestWorkbook w = tester.getWorkbook()) {
        try (TestWorkbook w = new TestWorkbook(outputFile)) {
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
