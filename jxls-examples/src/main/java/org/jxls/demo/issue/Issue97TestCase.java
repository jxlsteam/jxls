package org.jxls.demo.issue;

import org.jxls.common.Context;
import org.jxls.demo.guide.Employee;
import org.jxls.util.JxlsHelper;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * 'jx:each' command not processing Division formula having cell ref outside each
 */
public class Issue97TestCase {

// --------- SETTINGS ---------

    // define the lists which would be used in the template
    final static String INPUT_FILE_PATH = "issue97_template.xlsx";
    final static String OUTPUT_FILE_PATH = "target/issue97_output.xlsx";

    // --------- -------- ---------

    public static void main(String[] args) throws IOException, ParseException {
        try(InputStream is = Issue103TestCase.class.getResourceAsStream(INPUT_FILE_PATH)) {
            try (OutputStream os = new FileOutputStream(OUTPUT_FILE_PATH)) {
                Context context = new Context();
                List<Employee> employees = generateSampleEmployeeData("ACCOUNT");
                Map<String, List<Employee>> beans = new HashMap<>();
                beans.put("PEUS", generateSampleEmployeeData("PEUS"));
                beans.put("PEUSA", generateSampleEmployeeData("PEUSA"));
                beans.put("PEUSB", generateSampleEmployeeData("PEUSB"));
                context.putVar("employees", employees);
                context.putVar("map", beans);
                JxlsHelper.getInstance().processTemplate(is, os, context);
            }
        }
    }

    private static List<Employee> generateSampleEmployeeData(String buGroup) throws ParseException {
        List<Employee> employees = new ArrayList<Employee>();
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MMM-dd", Locale.US);
        employees.add(new Employee("Elsa", dateFormat.parse("1970-Jul-10"), 1500, 0.15, buGroup));
        employees.add(new Employee("Elsa", dateFormat.parse("1970-Jul-10"), 1500, 0.15, buGroup));
        employees.add(new Employee("Elsa", dateFormat.parse("1970-Jul-10"), 1500, 0.15, buGroup));
        employees.add(new Employee("Oleg", dateFormat.parse("1973-Apr-30"), 2300, 0.25, buGroup));
        employees.add(new Employee("Neil", dateFormat.parse("1975-Oct-05"), 2500, 0.00, buGroup));
        return employees;
    }


    public Employee generateSampleEmployee() throws ParseException {

        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MMM-dd", Locale.US);
        Employee employees = new Employee("Elsa", dateFormat.parse("1970-Jul-10"), 1500, 0.15, "PEUS");
        return employees;
    }
}
