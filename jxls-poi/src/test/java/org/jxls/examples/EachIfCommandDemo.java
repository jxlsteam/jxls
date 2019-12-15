package org.jxls.examples;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

import org.junit.Test;
import org.jxls.area.XlsArea;
import org.jxls.command.Command;
import org.jxls.command.EachCommand;
import org.jxls.command.IfCommand;
import org.jxls.common.AreaRef;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.entity.Department;
import org.jxls.transform.Transformer;
import org.jxls.util.TransformerFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * @author Leonid Vysochyn Date: 1/30/12 12:15 PM
 */
public class EachIfCommandDemo {
    private static final Logger logger = LoggerFactory.getLogger(EachIfCommandDemo.class);
    private static final String template = "each_if_demo.xls";
    private static final String output = "target/each_if_demo_output.xls";

    @Test
    public void test() throws IOException {
        logger.info("Running Each/If Commands demo");
        List<Department> departments = Department.createDepartments();
        logger.info("Opening input stream");
        try (InputStream is = EachIfCommandDemo.class.getResourceAsStream(template)) {
            try (OutputStream os = new FileOutputStream(output)) {
                Transformer transformer = TransformerFactory.createTransformer(is, os);
                // uncomment in case you want to avoid using ThreadLocals in ExpressionEvaluator
                // implementations
                // transformer.getTransformationConfig().setExpressionEvaluator(new
                // JexlExpressionEvaluatorNoThreadLocal());
                logger.info("Creating area");
                XlsArea xlsArea = new XlsArea("Template!A1:G15", transformer);
                XlsArea departmentArea = new XlsArea("Template!A2:G13", transformer);
                EachCommand departmentEachCommand = new EachCommand("department", "departments", departmentArea);
                XlsArea employeeArea = new XlsArea("Template!A9:F9", transformer);
                XlsArea ifArea = new XlsArea("Template!A18:F18", transformer);
                XlsArea elseArea = new XlsArea("Template!A9:F9", transformer);
                IfCommand ifCommand = new IfCommand("employee.payment <= 2000", ifArea, elseArea);
                employeeArea.addCommand(new AreaRef("Template!A9:F9"), ifCommand);
                Command employeeEachCommand = new EachCommand("employee", "department.staff", employeeArea);
                departmentArea.addCommand(new AreaRef("Template!A9:F9"), employeeEachCommand);
                xlsArea.addCommand(new AreaRef("Template!A2:F12"), departmentEachCommand);
                Context context = new Context();
                context.putVar("departments", departments);
                logger.info("Applying at cell " + new CellRef("Down!B2"));
                xlsArea.applyAt(new CellRef("Down!B2"), context);
                xlsArea.processFormulas();
                logger.info("Setting EachCommand direction to Right");
                departmentEachCommand.setDirection(EachCommand.Direction.RIGHT);
                logger.info("Applying at cell " + new CellRef("Right!A1"));
                xlsArea.reset();
                xlsArea.applyAt(new CellRef("Right!A1"), context);
                xlsArea.processFormulas();
                logger.info("Complete");
                transformer.write();
                logger.info("written to file");
            }
        }
    }
}
