package org.jxls.demo;

import org.jxls.demo.guide.Employee;
import org.jxls.template.SimpleExporter;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Locale;

/**
 * Created by Leonid Vysochyn on 19-Jul-15.
 */
public class SimpleExporterDemo {
    private static String template = "simple_export_template.xlsx";

    static Logger logger = LoggerFactory.getLogger(GridCommandDemo.class);

    public static void main(String[] args) throws ParseException, IOException {
        logger.info("Running Simple Export demo");
        try(OutputStream os1 = new FileOutputStream("target/simple_export_output1.xls")) {
            List<Employee> employees = generateSampleEmployeeData();
            List<String> headers = Arrays.asList("Name", "Birthday", "Payment");
            SimpleExporter exporter = new SimpleExporter();
            exporter.gridExport(headers, employees, "name, birthDate, payment", os1);

            // now let's show how to register custom template
            try (InputStream is = SimpleExporterDemo.class.getResourceAsStream(template)) {
                try (OutputStream os2 = new FileOutputStream("target/simple_export_output2.xlsx")) {
                    exporter.registerGridTemplate(is);
                    headers = Arrays.asList("Name", "Payment", "Birth Date");
                    exporter.gridExport(headers, employees, "name,payment, birthDate,", os2);
                }
            }
        }
    }

    private static List<Employee> generateSampleEmployeeData() throws ParseException {
        List<Employee> employees = new ArrayList<Employee>();
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MMM-dd", Locale.US);
        employees.add( new Employee("Elsa", dateFormat.parse("1970-Jul-10"), 1500, 0.15) );
        employees.add(new Employee("Oleg", dateFormat.parse("1973-Apr-30"), 2300, 0.25));
        employees.add(new Employee("Neil", dateFormat.parse("1975-Oct-05"), 2500, 0.00));
        employees.add(new Employee("Maria", dateFormat.parse("1978-Jan-07"), 1700, 0.15));
        employees.add(new Employee("John", dateFormat.parse("1969-May-30"), 2800, 0.20));
        return employees;
    }
}
