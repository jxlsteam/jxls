package org.jxls.examples;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

import org.junit.Test;
import org.jxls.area.Area;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.command.EachCommand;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.entity.Department;
import org.jxls.transform.Transformer;
import org.jxls.transform.poi.PoiTransformer;
import org.jxls.util.TransformerFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * @author Leonid Vysochyn
 */
public class XlsCommentBuilderDemo {
    private static final Logger logger = LoggerFactory.getLogger(XlsCommentBuilderDemo.class);
    private static final String template = "comment_markup_demo.xls";
    private static final String output = "target/comment_builder_output.xls";

    @Test
    public void test() throws IOException {
        logger.info("Running XLS Comment builder demo");
        List<Department> departments = Department.createDepartments();
        logger.info("Opening input stream");
        try (InputStream is = XlsCommentBuilderDemo.class.getResourceAsStream(template)) {
            try (OutputStream os = new FileOutputStream(output)) {
                Transformer transformer = TransformerFactory.createTransformer(is, os);
                AreaBuilder areaBuilder = new XlsCommentAreaBuilder();
                List<Area> xlsAreaList = areaBuilder.build(transformer, false);
                Area xlsArea = xlsAreaList.get(0);
                Context context = PoiTransformer.createInitialContext();
                context.putVar("departments", departments);
                logger.info("Applying area " + xlsArea.getAreaRef() + " at cell " + new CellRef("Down!A1"));
                xlsArea.applyAt(new CellRef("Down!A1"), context);
                xlsArea.processFormulas();
                xlsArea.reset();
                EachCommand eachCommand = (EachCommand) xlsArea.findCommandByName("each").get(0);
                eachCommand.setDirection(EachCommand.Direction.RIGHT);
                logger.info("Applying area " + xlsArea.getAreaRef() + " at cell " + new CellRef("Right!A1"));
                xlsArea.applyAt(new CellRef("Right!A1"), context);
                xlsArea.processFormulas();
                logger.info("Complete");
                transformer.write();
                logger.info("written to file");
            }
        }
    }
}
