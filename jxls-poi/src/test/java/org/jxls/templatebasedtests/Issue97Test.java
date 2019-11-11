package org.jxls.templatebasedtests;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.common.Context;
import org.jxls.entity.Employee;

/**
 * 'jx:each' command not processing Division formula having cell ref outside each
 */
public class Issue97Test {

    @Test
    public void test() throws Exception {
        Context context = new Context();
        context.putVar("map", getBeans());
        context.putVar("employees", generateSampleEmployeeData("ACCOUNT"));

        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);
        
        // TODO assertions
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
