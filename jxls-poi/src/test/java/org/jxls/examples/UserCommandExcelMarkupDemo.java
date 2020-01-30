package org.jxls.examples;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.Test;
import org.jxls.area.Area;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.entity.Employee;
import org.jxls.examples.UserCommandDemo.GroupRowCommand;
import org.jxls.transform.Transformer;
import org.jxls.util.TransformerFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * @author Leonid Vysochyn
 */
public class UserCommandExcelMarkupDemo {
    private static final Logger logger = LoggerFactory.getLogger(UserCommandExcelMarkupDemo.class);
    private static final String template = "user_command_markup_template.xls";
    private static final String output = "target/user_command_markup_output.xls";

    @Test
    public void test() throws IOException, InvalidFormatException, ParseException {
        logger.info("Running User Command Markup demo");
        List<Employee> employees = Employee.generateSampleEmployeeData();
        logger.info("Opening input stream");
        try (InputStream is = UserCommandExcelMarkupDemo.class.getResourceAsStream(template)) {
            try (OutputStream os = new FileOutputStream(output)) {
                Transformer transformer = TransformerFactory.createTransformer(is, os);
                AreaBuilder areaBuilder = new XlsCommentAreaBuilder(transformer);
                XlsCommentAreaBuilder.addCommandMapping("groupRow", GroupRowCommand.class);
                List<Area> xlsAreaList = areaBuilder.build();
                Area xlsArea = xlsAreaList.get(0);
                Context context = new Context();
                context.putVar("employees", employees);
                xlsArea.applyAt(new CellRef("Result!A1"), context);
                transformer.write();
                logger.info("Finished UserCommandExcelMarkupDemo");
            }
        }
    }
}
