package org.jxls.examples;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.util.List;

import org.junit.Test;
import org.jxls.area.XlsArea;
import org.jxls.command.EachCommand;
import org.jxls.command.IfCommand;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.entity.Employee;
import org.jxls.transform.Transformer;
import org.jxls.util.TransformerFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * @author Leonid Vysochyn
 *         Date: 12/30/13
 */
public class NestedCommandJavaAPIDemo {
    private static final Logger logger = LoggerFactory.getLogger(NestedCommandJavaAPIDemo.class);

    @Test
    public void test() throws ParseException, IOException {
        logger.info("Running Nested Command JavaAPI demo");
        List<Employee> employees = Employee.generateSampleEmployeeData();
        try (InputStream is = NestedCommandJavaAPIDemo.class.getResourceAsStream("nested_command_javaapi_template.xls")) {
            try (OutputStream os = new FileOutputStream("target/nested_command_javaapi_output.xls")) {
                Transformer transformer = TransformerFactory.createTransformer(is, os);
                XlsArea xlsArea = new XlsArea("Template!A1:D4", transformer);
                XlsArea employeeArea = new XlsArea("Template!A4:D4", transformer);
                EachCommand employeeEachCommand = new EachCommand("employee", "employees", employeeArea);
                xlsArea.addCommand("A4:D4", employeeEachCommand);
                XlsArea ifArea = new XlsArea("Template!A6:D6", transformer);
                XlsArea elseArea = new XlsArea("Template!A4:D4", transformer);
                IfCommand ifCommand = new IfCommand("employee.payment <= 2000", ifArea, elseArea);
                employeeArea.addCommand("Template!A4:D4", ifCommand);
                Context context = new Context();
                context.putVar("employees", employees);
                xlsArea.applyAt(new CellRef("Result!A1"), context);
                transformer.write();
            }
        }
    }
}
