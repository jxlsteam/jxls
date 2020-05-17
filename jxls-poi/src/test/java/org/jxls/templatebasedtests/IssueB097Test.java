package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;
import org.jxls.entity.Employee;

/**
 * 'jx:each' command not processing Division formula having cell ref outside each
 */
public class IssueB097Test {

    @Test
    public void test() throws Exception {
        // Prepare
        Context context = new Context();
        context.putVar("map", getBeans());
        context.putVar("employees", generateSampleEmployeeData("ACCOUNT"));

        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet(0);
            assertEquals(0.161290322580645d, w.getCellValueAsDouble(19, 5), 0.000000000000001d);
            assertEquals(0.161290322580645d, w.getCellValueAsDouble(20, 5), 0.000000000000001d);
            assertEquals(0.161290322580645d, w.getCellValueAsDouble(21, 5), 0.000000000000001d);
            assertEquals(0.247311827956989d, w.getCellValueAsDouble(22, 5), 0.000000000000001d);
            assertEquals(0.268817204301075d, w.getCellValueAsDouble(23, 5), 0.000000000000001d);
            assertEquals(0.333333333333333d, w.getCellValueAsDouble(24, 5), 0.000000000000001d);
        }
    }

    private Map<String, List<Employee>> getBeans() throws ParseException {
        Map<String, List<Employee>> beans = new HashMap<>();
        beans.put("PEUS", generateSampleEmployeeData("PEUS"));
        beans.put("PEUSA", generateSampleEmployeeData("PEUSA"));
        beans.put("PEUSB", generateSampleEmployeeData("PEUSB"));
        return beans;
    }

    private static List<Employee> generateSampleEmployeeData(String buGroup) throws ParseException {
        List<Employee> employees = new ArrayList<>();
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MMM-dd", Locale.US);
        employees.add(new Employee("Elsa", dateFormat.parse("1970-Jul-10"), 1500, 0.15, buGroup));
        employees.add(new Employee("Elsa", dateFormat.parse("1970-Jul-10"), 1500, 0.15, buGroup));
        employees.add(new Employee("Elsa", dateFormat.parse("1970-Jul-10"), 1500, 0.15, buGroup));
        employees.add(new Employee("Oleg", dateFormat.parse("1973-Apr-30"), 2300, 0.25, buGroup));
        employees.add(new Employee("Neil", dateFormat.parse("1975-Oct-05"), 2500, 0.00, buGroup));
        return employees;
    }
}
