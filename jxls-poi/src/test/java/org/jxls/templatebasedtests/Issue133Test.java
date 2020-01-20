package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;

import java.util.ArrayList;
import java.util.List;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.command.TestEmployee;
import org.jxls.common.Context;

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
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
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
