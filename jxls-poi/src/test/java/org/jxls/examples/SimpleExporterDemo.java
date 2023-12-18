package org.jxls.examples;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.util.Arrays;
import java.util.List;

import org.junit.Test;
import org.jxls.entity.Employee;
import org.jxls.template.SimpleExporter;

/**
 * Created by Leonid Vysochyn on 19-Jul-15.
 */
public class SimpleExporterDemo {
    private static final String template = "simple_export_template.xlsx";

    @Test
    public void test() throws ParseException, IOException {
        try(OutputStream os1 = new FileOutputStream("target/simple_export_output1.xls")) {
            List<Employee> employees = Employee.generateSampleEmployeeData();
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
}
