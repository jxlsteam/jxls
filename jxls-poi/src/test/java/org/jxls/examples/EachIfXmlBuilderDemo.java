package org.jxls.examples;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

import org.junit.Test;
import org.jxls.area.Area;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xml.XmlAreaBuilder;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.entity.Department;
import org.jxls.transform.Transformer;
import org.jxls.util.TransformerFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * @author Leonid Vysochyn
 *         Date: 2/14/12 3:59 PM
 */
public class EachIfXmlBuilderDemo {
    private static final Logger logger = LoggerFactory.getLogger(EachIfCommandDemo.class);
    private static final String template = "each_if_demo.xls";
    private static final String xmlConfig = "each_if_demo.xml";
    private static final String output = "target/each_if_xml_builder_output.xls";

    @Test
    public void test() throws IOException {
        logger.info("Running Each/If Command XML Builder demo");
        execute();
    }

    public static void execute() throws IOException {
        List<Department> departments = Department.createDepartments();
        logger.info("Opening input stream");
        try(InputStream is = EachIfCommandDemo.class.getResourceAsStream(template)) {
            try (OutputStream os = new FileOutputStream(output)) {
                Transformer transformer = TransformerFactory.createTransformer(is, os);
                logger.info("Creating areas");
                InputStream configInputStream = EachIfXmlBuilderDemo.class.getResourceAsStream(xmlConfig);
                AreaBuilder areaBuilder = new XmlAreaBuilder(configInputStream, transformer);
                List<Area> xlsAreaList = areaBuilder.build();
                Area xlsArea = xlsAreaList.get(0);
                Area xlsArea2 = xlsAreaList.get(1);
                Context context = new Context();
                context.putVar("departments", departments);
                logger.info("Applying first area at cell " + new CellRef("Down!A1"));
                xlsArea.applyAt(new CellRef("Down!A1"), context);
                xlsArea.processFormulas();
                logger.info("Applying second area at cell " + new CellRef("Right!A1"));
                xlsArea.reset();
                xlsArea2.applyAt(new CellRef("Right!A1"), context);
                xlsArea2.processFormulas();
                logger.info("Complete");
                transformer.write();
                logger.info("written to file");
            }
        }
    }
}
