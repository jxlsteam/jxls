package org.jxls.examples;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.util.Arrays;
import java.util.List;

import org.junit.Test;
import org.jxls.common.Context;
import org.jxls.entity.Employee;
import org.jxls.util.JxlsHelper;

// old filename: MultiSheetDemo.java
public class MultiSheetMarkup3Demo {

    @Test
    public void test() throws ParseException, IOException {
        List<Employee> employees = Employee.generateSampleEmployeeData();
        try (InputStream is = MultiSheetMarkup3Demo.class.getResourceAsStream("multisheet_template.xls")) {
            try (OutputStream os = new FileOutputStream("target/multisheet_output.xls")) {
                Context context = new Context();
                context.putVar("employees", employees);
                context.putVar("sheetNames", Arrays.asList("Emp 1", "Emp 2", "Emp 3", "Emp 4", "Emp 5"));
                JxlsHelper.getInstance().processTemplate(is, os, context);
            }
        }
    }
}
