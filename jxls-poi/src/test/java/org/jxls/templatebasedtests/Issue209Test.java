package org.jxls.templatebasedtests;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.junit.Assert;
import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;

/**
 * It should be examined here whether the simultaneous use of the jx:each attributes
 * groupBy and select has been implemented in a meaningful way.
 */
public class Issue209Test {

    @Test
    public void test() {
        // Prepare
        List<Map<String, String>> employees = new ArrayList<>();
        employees.add(createEmployee("Department A", "Claudia", "Amsterdam")); // First must not be Geldern!
        employees.add(createEmployee("Department A", "Dagmar", "Geldern"));
        employees.add(createEmployee("Department A", "Sven", "Geldern"));
        employees.add(createEmployee("Department B", "Doris", "Wetten"));
        employees.add(createEmployee("Department B", "Melanie", "Geldern"));
        employees.add(createEmployee("Department C", "Stefan", "Bruegge"));
        Context context = new Context();
        context.putVar("employees", employees);
        
        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet(0);
            Assert.assertEquals("Department A", w.getCellValueAsString(3, 1));
            Assert.assertEquals("Dagmar", w.getCellValueAsString(4, 2));
            Assert.assertEquals("Sven", w.getCellValueAsString(5, 2));
            Assert.assertEquals("Department B", w.getCellValueAsString(7, 1));
            Assert.assertEquals("Melanie", w.getCellValueAsString(8, 2));
            Assert.assertEquals("Geldern", w.getCellValueAsString(8, 3));
        }
    }
    
    private Map<String, String> createEmployee(String department, String name, String city) {
        Map<String, String> map = new HashMap<>();
        map.put("department", department);
        map.put("name", name);
        map.put("city", city);
        return map;
    }
}
