package org.jxls.demo.issue;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.util.List;

import org.jxls.common.Context;
import org.jxls.demo.guide.Employee;
import org.jxls.demo.guide.ObjectCollectionDemo;
import org.jxls.util.JxlsHelper;

/**
 * Conditional formatting copying issue
 */
public class Issue89TestCase {
    private final static String INPUT_FILE_PATH = "issue89_template.xlsx";
    private final static String OUTPUT_FILE_PATH = "target/issue89_output.xlsx";

    public static void main(String[] args) throws IOException, ParseException {
        try (InputStream is = Issue89TestCase.class.getResourceAsStream(INPUT_FILE_PATH)) {
            try (OutputStream os = new FileOutputStream(OUTPUT_FILE_PATH)) {
                List<Employee> employees = ObjectCollectionDemo.generateSampleEmployeeData();
                Context context = new Context();
                context.putVar("employees", employees);
                JxlsHelper.getInstance().processTemplate(is, os, context);
            }
        }
    }
}
