package org.jxls.examples;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.area.Area;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.ContextImpl;
import org.jxls.entity.Employee;
import org.jxls.transform.Transformer;

/**
 * @author Leonid Vysochyn
 */
public class UserCommandExcelMarkupDemo {
    private static final String template = "user_command_markup_template.xls";
    private static final String output = "target/user_command_markup_output.xls";

    @Test
    public void test() throws IOException, InvalidFormatException, ParseException {
        List<Employee> employees = Employee.generateSampleEmployeeData();
        try (InputStream is = UserCommandExcelMarkupDemo.class.getResourceAsStream(template)) {
            try (OutputStream os = new FileOutputStream(output)) {
                Transformer transformer = Jxls3Tester.createTransformer(is, os);
                XlsCommentAreaBuilder areaBuilder = new XlsCommentAreaBuilder();
                areaBuilder.addCommandMapping("groupRow", GroupRowCommand.class);
                List<Area> xlsAreaList = areaBuilder.build(transformer, true);
                Area xlsArea = xlsAreaList.get(0);
                Context context = new ContextImpl();
                context.putVar("employees", employees);
                xlsArea.applyAt(new CellRef("Result!A1"), context);
                transformer.write();
            }
        }
    }
}
