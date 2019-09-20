package org.jxls.demo.issue;

import org.jxls.common.Context;
import org.jxls.demo.guide.Employee;
import org.jxls.util.JxlsHelper;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Date;

/**
 * Wrong average on 2nd sheet
 */
public class Issue167TestCase {

// --------- SETTINGS ---------

    // define the lists which would be used in the template
    final static String INPUT_FILE_PATH = "issue167_Template.xls";
    final static String OUTPUT_FILE_PATH = "target/issue167_Output.xls";

    // --------- -------- ---------

    public static void main(String[] args) throws IOException {
        try(InputStream is = Issue103TestCase.class.getResourceAsStream(INPUT_FILE_PATH)) {
            try (OutputStream os = new FileOutputStream(OUTPUT_FILE_PATH)) {
                Context context = new Context();
                context.putVar("var1", "Variable 1");
                context.putVar("var2", new Employee("Leo", new Date(), 1000, 0.10));
                JxlsHelper.getInstance().processTemplate(is, os, context);
            }
        }
    }}
