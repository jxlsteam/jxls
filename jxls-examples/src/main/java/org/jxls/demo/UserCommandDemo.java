package org.jxls.demo;

import org.jxls.area.Area;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xml.XmlAreaBuilder;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.demo.model.Department;
import org.jxls.transform.Transformer;
import org.jxls.util.TransformerFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

/**
 * @author Leonid Vysochyn
 *         Date: 2/21/12 4:30 PM
 */
public class UserCommandDemo {
    static Logger logger = LoggerFactory.getLogger(UserCommandDemo.class);
    private static String template = "each_if_demo.xls";
    private static String xmlConfig = "user_command_demo.xml";
    private static String output = "target/user_command_output.xls";

    public static void main(String[] args) throws IOException {
        logger.info("Running User Command demo");
        execute();
    }

    public static void execute() throws IOException {
        List<Department> departments = EachIfCommandDemo.createDepartments();
        logger.info("Opening input stream");
        try(InputStream is = EachIfCommandDemo.class.getResourceAsStream(template)) {
            try (OutputStream os = new FileOutputStream(output)) {
                Transformer transformer = TransformerFactory.createTransformer(is, os);
                System.out.println("Creating areas");
                try (InputStream configInputStream = UserCommandDemo.class.getResourceAsStream(xmlConfig)) {
                    AreaBuilder areaBuilder = new XmlAreaBuilder(configInputStream, transformer);
                    List<Area> xlsAreaList = areaBuilder.build();
                    Area xlsArea = xlsAreaList.get(0);
                    Context context = new Context();
                    context.putVar("departments", departments);
                    logger.info("Applying area at cell " + new CellRef("Down!A1"));
                    xlsArea.applyAt(new CellRef("Down!A1"), context);
                    xlsArea.processFormulas();
                    logger.info("Complete");
                    transformer.write();
                    logger.info("Written to file");
                }
            }
        }
    }

}
