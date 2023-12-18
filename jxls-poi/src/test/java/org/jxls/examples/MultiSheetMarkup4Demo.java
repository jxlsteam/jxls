package org.jxls.examples;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

import org.junit.Test;
import org.jxls.area.XlsArea;
import org.jxls.command.CellRefGenerator;
import org.jxls.command.Command;
import org.jxls.command.EachCommand;
import org.jxls.command.IfCommand;
import org.jxls.common.AreaRef;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.entity.Department;
import org.jxls.logging.JxlsLogger;
import org.jxls.transform.Transformer;
import org.jxls.util.TransformerFactory;

// old filename: MultipleSheetDemo.java
public class MultiSheetMarkup4Demo {
    private static final String template = "each_if_demo.xls";
    private static final String output = "target/multiple_sheet_demo_output.xls";

    @Test
    public void test() throws IOException {
        List<Department> departments = Department.createDepartments();
        try (InputStream is = EachIfCommandDemo.class.getResourceAsStream(template)) {
            try (OutputStream os = new FileOutputStream(output)) {
                Transformer transformer = TransformerFactory.createTransformer(is, os);
                XlsArea xlsArea = new XlsArea("Template!A1:G15", transformer);
                XlsArea departmentArea = new XlsArea("Template!A2:G12", transformer);
                EachCommand departmentEachCommand = new EachCommand("department", "departments", departmentArea, new SimpleCellRefGenerator());
                XlsArea employeeArea = new XlsArea("Template!A9:F9", transformer);
                XlsArea ifArea = new XlsArea("Template!A18:F18", transformer);
                IfCommand ifCommand = new IfCommand("employee.payment <= 2000", ifArea, new XlsArea("Template!A9:F9", transformer));
                employeeArea.addCommand(new AreaRef("Template!A9:F9"), ifCommand);
                Command employeeEachCommand = new EachCommand("employee", "department.staff", employeeArea);
                departmentArea.addCommand(new AreaRef("Template!A9:F9"), employeeEachCommand);
                xlsArea.addCommand(new AreaRef("Template!A2:F12"), departmentEachCommand);
                Context context = new Context();
                context.putVar("departments", departments);
                xlsArea.applyAt(new CellRef("Sheet!A1"), context);
                xlsArea.processFormulas();
                transformer.write();
            }
        }
    }

    public static class SimpleCellRefGenerator implements CellRefGenerator {

        @Override
        public CellRef generateCellRef(int index, Context context, JxlsLogger logger) {
            return new CellRef("sheet" + index + "!B2");
        }
    }
}
