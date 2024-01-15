package org.jxls.examples;

import static org.junit.Assert.assertEquals;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.util.Arrays;
import java.util.List;

import org.junit.Test;
import org.jxls.TestWorkbook;
import org.jxls.entity.Employee;
import org.jxls.template.SimpleExporter;

public class SimpleExporterDemo {
    
    @Test
    public void test() throws ParseException, IOException {
        // Prepare
        List<String> headers;
        List<Employee> employees = Employee.generateSampleEmployeeData();
        SimpleExporter exporter = new SimpleExporter();

        // Test
        headers = Arrays.asList("Name", "Birthday", "Payment");
        String out1 = "target/simple_export_output1.xls";
        try (OutputStream os1 = new FileOutputStream(out1)) {
            exporter.gridExport(headers, employees, "name, birthDate, payment", os1);
        }
        
        // Verify
        try (TestWorkbook w = new TestWorkbook(new File(out1))) {
            w.selectSheet(0);
            assertEquals(Double.valueOf(1500), w.getCellValueAsDouble(2, 3), 0.005d);
        }

        // Test: now let's show how to register custom template
        String out2 = "target/simple_export_output2.xlsx";
        try (InputStream is = SimpleExporterDemo.class.getResourceAsStream("simple_export_template.xlsx")) {
            try (OutputStream os2 = new FileOutputStream(out2)) {
                exporter.registerGridTemplate(is);
                headers = Arrays.asList("Name", "Payment", "Birth Date");
                exporter.gridExport(headers, employees, "name,payment, birthDate,", os2);
            }
        }
        
        // Verify
        try (TestWorkbook w = new TestWorkbook(new File(out2))) {
            w.selectSheet(0);
            assertEquals(Double.valueOf(1500), w.getCellValueAsDouble(2, 2), 0.005d);
        }
    }
}
