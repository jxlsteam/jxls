package org.jxls.examples;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.util.List;

import org.junit.Test;
import org.jxls.common.Context;
import org.jxls.entity.Employee;
import org.jxls.util.JxlsHelper;

/**
 * 2nd Highlight Demo
 * 
 * @author Leonid Vysochyn Date: 10/22/13
 */
public class Highlight2Demo {

    @Test
    public void test() throws ParseException, IOException {
        List<Employee> employees = Employee.generateSampleEmployeeData();
        try (InputStream is = ObjectCollectionDemo.class.getResourceAsStream("highlight_template.xls")) {
            try (OutputStream os = new FileOutputStream("target/highlight_output.xls")) {
                Context context = new Context();
                context.putVar("employees", employees);
                JxlsHelper.getInstance().processTemplateAtCell(is, os, context, "Result!A1");
            }
        }
    }
}
