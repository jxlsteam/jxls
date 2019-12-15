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
 * @author Leonid Vysochyn
 */
public class ParameterizedFormulasDemo {
    private static final Logger logger = LoggerFactory.getLogger(ParameterizedFormulasDemo.class);

    @Test
    public void test() throws ParseException, IOException {
        logger.info("Running Parameterized Formulas demo");
        List<Employee> employees = Employee.generateSampleEmployeeData();
        try(InputStream is = ParameterizedFormulasDemo.class.getResourceAsStream("param_formulas_template.xls")) {
            try (OutputStream os = new FileOutputStream("target/param_formulas_output.xls")) {
                Context context = new Context();
                context.putVar("employees", employees);
                context.putVar("bonus", 0.1);
                JxlsHelper.getInstance().processTemplateAtCell(is, os, context, "Result!A1");
            }
        }
    }
}
