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
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Created by Leonid Vysochyn on 22-Jul-15.
 */
public class CustomExpressionNotationDemo {
    private static final Logger logger = LoggerFactory.getLogger(CustomExpressionNotationDemo.class);
    private static final String TEMPLATE = "custom_expression_notation_template.xlsx";
    private static final String OUTPUT = "target/custom_expression_notation_output.xlsx";

    @Test
    public void test() throws ParseException, IOException {
        logger.info("Running Custom Expression Notation demo");
        List<Employee> employees = Employee.generateSampleEmployeeData();
        try(InputStream is = CustomExpressionNotationDemo.class.getResourceAsStream(TEMPLATE)) {
            try (OutputStream os = new FileOutputStream(OUTPUT)) {
                Context context = new Context();
                context.putVar("employees", employees);
                JxlsHelper.getInstance().buildExpressionNotation("[[", "]]").processTemplate(is, os, context);
            }
        }
    }
}
