package org.jxls.command;

import static org.junit.Assert.assertEquals;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.junit.Test;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

/**
 * Group by nested property
 */
public class Issue133Test {

    @Test
    public void groupingWithNestedGroupKey() throws Exception {
        // Prepare
        Context context = new Context();
        context.putVar("employees", createEmployees());
        
        // Test
/*TODO  JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);
*/
        File outputFile = new File("target/Issue133Test_output.xlsx");
        try (InputStream in = getClass().getResourceAsStream("Issue133Test.xlsx")) {
            try (FileOutputStream out = new FileOutputStream(outputFile)) {
                JxlsHelper.getInstance().processTemplate(in, out, context);
            }
        }
        
        // Verify
//TODO  try (TestWorkbook w = tester.getWorkbook()) {
        try (TestWorkbook w = new TestWorkbook(outputFile)) {
            w.selectSheet("GroupByNestedProperty");
            assertEquals("Mayor", w.getCellValueAsString(2, 1)); 
            assertEquals("Finance", w.getCellValueAsString(3, 1)); 
            assertEquals("Audit office", w.getCellValueAsString(4, 1)); 
        }
    }

    // also used by OrderByTest
    public static List<TestEmployee> createEmployees() {
        List<TestEmployee> employees = new ArrayList<>();
        employees.add(new TestEmployee("Mayor", "Sven", "", "", 0).withDepartmentKey("01"));
        employees.add(new TestEmployee("Finance", "Thomas", "", "", 0).withDepartmentKey("03"));
        employees.add(new TestEmployee("Mayor", "Herbert", "", "", 0).withDepartmentKey("01"));
        employees.add(new TestEmployee("Audit office", "Markus", "", "", 0).withDepartmentKey("03-1"));
        return employees;
    }
}
