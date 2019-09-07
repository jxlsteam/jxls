package org.jxls.demo.issue;

import org.jxls.common.Context;
import org.jxls.demo.guide.Employee;
import org.jxls.demo.guide.ObjectCollectionDemo;
import org.jxls.transform.Transformer;
import org.jxls.util.JxlsHelper;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.util.Arrays;
import java.util.List;

public class Issue153TestCase
{
    private final static String INPUT_FILE_PATH = "issue153_template.xlsx";

    private final static String OUTPUT_FILE_PATH = "target/issue153_output.xlsx";

    public static void main(String[] args) throws IOException, ParseException {
        try (InputStream is = Issue153TestCase.class.getResourceAsStream(INPUT_FILE_PATH))
        {
            try (OutputStream os = new FileOutputStream(OUTPUT_FILE_PATH))
            {
                List<Employee> employees = ObjectCollectionDemo.generateSampleEmployeeData();
                Context context = new Context();
                context.putVar("employees", employees);
                JxlsHelper jxlsHelper = JxlsHelper.getInstance();
                Transformer transformer = jxlsHelper.createTransformer(is, os);
                jxlsHelper.processTemplate(context, transformer);
            }
        }
    }

}