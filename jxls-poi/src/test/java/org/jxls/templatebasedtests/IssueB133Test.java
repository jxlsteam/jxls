package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;

import java.util.ArrayList;
import java.util.List;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;
import org.jxls.entity.Employee;

/**
 * Group by nested property
 */
public class IssueB133Test {

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
    public static List<Employee> createEmployees() {
        List<Employee> employees = new ArrayList<>();
        employees.add(createEmployee("Mayor", "Sven").withDepartmentKey("01"));
        employees.add(createEmployee("Finance", "Thomas").withDepartmentKey("03"));
        employees.add(createEmployee("Mayor", "Herbert").withDepartmentKey("01"));
        employees.add(createEmployee("Audit office", "Markus").withDepartmentKey("03-1"));
        return employees;
    }
    
    private static Employee createEmployee(String department, String name) {
        Employee employee = new Employee(name, null, 0, 0);
        employee.setBuGroup(department);
        return employee;
    }
}
