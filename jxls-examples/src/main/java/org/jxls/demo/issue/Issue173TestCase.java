package org.jxls.demo.issue;

import org.jxls.common.Context;
import org.jxls.demo.guide.Employee;
import org.jxls.demo.guide.ObjectCollectionDemo;
import org.jxls.util.JxlsHelper;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.util.List;

public class Issue173TestCase {

    public static void main(String[] args) throws IOException, ParseException {

        try(InputStream is = Issue173TestCase.class.getResourceAsStream("issue173_template.xls")) {
            try (OutputStream os = new FileOutputStream("target/issue173_output.xls")) {
                List<Employee> employees = ObjectCollectionDemo.generateSampleEmployeeData();
                Context context = new Context();
                context.putVar("employees", employees);
                JxlsHelper.getInstance().processTemplate(is, os, context);
            }
        }
    }
}
