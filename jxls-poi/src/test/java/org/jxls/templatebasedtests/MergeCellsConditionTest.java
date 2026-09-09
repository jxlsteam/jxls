package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import java.util.ArrayList;
import java.util.List;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;
import org.jxls.entity.Employee;

/**
 * Verifies jx:mergeCells optional condition attribute:
 * merge only when condition is true (e.g. group size &gt; 1).
 * 
 * @see <a href="https://github.com/jxlsteam/jxls/issues/427">issue #427</a>
 */
public class MergeCellsConditionTest {

    @Test
    public void conditionSkipsMergeForSingleRowGroups() {
        Context context = new Context();
        context.putVar("employees", createEmployees());

        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);

        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet(0);
            String merged = w.getMergedCells();
            // Finance (1 row at A2) must not merge; multi-row groups must merge column A
            assertFalse("Finance single row must not be merged: " + merged, merged.contains("A2:"));
            assertTrue("expected HR merge in " + merged, merged.contains("A3:A4"));
            assertTrue("expected IT merge in " + merged, merged.contains("A5:A6"));
            assertTrue("expected Sales merge in " + merged, merged.contains("A7:A9"));

            assertEquals("Finance", w.getCellValueAsString(2, 1));
            assertEquals("Carol", w.getCellValueAsString(2, 2));
            assertEquals("HR", w.getCellValueAsString(3, 1));
            assertEquals("Anna", w.getCellValueAsString(3, 2));
            assertEquals("Bob", w.getCellValueAsString(4, 2));
            assertEquals("IT", w.getCellValueAsString(5, 1));
            assertEquals("Sales", w.getCellValueAsString(7, 1));
        }
    }

    static List<Employee> createEmployees() {
        List<Employee> employees = new ArrayList<>();
        employees.add(emp("Carol", "Finance"));
        employees.add(emp("Anna", "HR"));
        employees.add(emp("Bob", "HR"));
        employees.add(emp("Elsa", "IT"));
        employees.add(emp("Oleg", "IT"));
        employees.add(emp("John", "Sales"));
        employees.add(emp("Maria", "Sales"));
        employees.add(emp("Neil", "Sales"));
        return employees;
    }

    private static Employee emp(String name, String dept) {
        Employee e = new Employee(name, null, 0, 0);
        e.setBuGroup(dept);
        return e;
    }
}
