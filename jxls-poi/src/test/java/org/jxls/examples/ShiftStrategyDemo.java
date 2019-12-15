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
 * Created by Leonid Vysochyn on 08-Aug-15.
 */
public class ShiftStrategyDemo {
    private static final Logger logger = LoggerFactory.getLogger(ShiftStrategyDemo.class);

    @Test
    public void test() throws ParseException, IOException {
        logger.info("Running Shift Strategy Demo");
        List<Employee> employees = Employee.generateSampleEmployeeData();
        try (InputStream is = ShiftStrategyDemo.class.getResourceAsStream("shiftstrategy_template.xlsx")) {
            try (OutputStream os = new FileOutputStream("target/shiftstrategy_output.xlsx")) {
                Context context = new Context();
                context.putVar("employees", employees);
                JxlsHelper.getInstance().processTemplateAtCell(is, os, context, "Result!A1");
            }
        }
    }
}
