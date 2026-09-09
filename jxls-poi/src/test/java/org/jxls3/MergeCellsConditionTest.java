package org.jxls3;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.junit.Assert;
import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.TestWorkbook;
import org.jxls.entity.Employee;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

/**
 * Verifies jx:mergeCells optional condition attribute:
 * merge only when condition is true (e.g. group size &gt; 1).
 *
 * @see <a href="https://github.com/jxlsteam/jxls/issues/427">issue #427</a>
 */
public class MergeCellsConditionTest {

    @Test
    public void conditionSkipsMergeForSingleRowGroups() {
        Map<String, Object> data = new HashMap<>();
        data.put("employees", createEmployees());

        Jxls3Tester tester = Jxls3Tester.xlsx(getClass());
        tester.test(data, JxlsPoiTemplateFillerBuilder.newInstance());

        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet(0);
            String merged = w.getMergedCells();
            Assert.assertFalse("Finance single row must not be merged: " + merged, merged.contains("A2:"));
            Assert.assertTrue("expected HR merge in " + merged, merged.contains("A3:A4"));
            Assert.assertTrue("expected IT merge in " + merged, merged.contains("A5:A6"));
            Assert.assertTrue("expected Sales merge in " + merged, merged.contains("A7:A9"));

            Assert.assertEquals("Finance", w.getCellValueAsString(2, 1));
            Assert.assertEquals("Carol", w.getCellValueAsString(2, 2));
            Assert.assertEquals("HR", w.getCellValueAsString(3, 1));
            Assert.assertEquals("Anna", w.getCellValueAsString(3, 2));
            Assert.assertEquals("Bob", w.getCellValueAsString(4, 2));
            Assert.assertEquals("IT", w.getCellValueAsString(5, 1));
            Assert.assertEquals("Sales", w.getCellValueAsString(7, 1));
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
