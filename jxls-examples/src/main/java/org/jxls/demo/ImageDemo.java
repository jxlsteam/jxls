package org.jxls.demo;

import org.jxls.area.XlsArea;
import org.jxls.command.ImageCommand;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.ImageType;
import org.jxls.demo.model.Department;
import org.jxls.transform.Transformer;
import org.jxls.util.JxlsHelper;
import org.jxls.util.TransformerFactory;
import org.jxls.util.Util;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

/**
 * @author Leonid Vysochyn
 */
public class ImageDemo {
    static Logger logger = LoggerFactory.getLogger(ImageDemo.class);
    private static String template = "image_demo.xlsx";
    private static String template2 = "image_demo2.xlsx";
    private static String output = "target/image_output.xlsx";
    private static String output2 = "target/image_output2.xlsx";
    private static String output3 = "target/image_output3.xlsx";

    public static void main(String[] args) throws IOException {
        logger.info("Running Image demo");
        execute();
        execute2();
        executeWithNullBytes();
    }

    public static void execute() throws IOException {
        logger.info("Opening input stream");
        try(InputStream is = ImageDemo.class.getResourceAsStream(template)) {
            try (OutputStream os = new FileOutputStream(output)) {
                Transformer transformer = TransformerFactory.createTransformer(is, os);
                XlsArea xlsArea = new XlsArea("Sheet1!A1:N30", transformer);
                Context context = new Context();
                InputStream imageInputStream = ImageDemo.class.getResourceAsStream("business.png");
                byte[] imageBytes = Util.toByteArray(imageInputStream);
                context.putVar("image", imageBytes);
                XlsArea imgArea = new XlsArea("Sheet1!A5:D15", transformer);
                xlsArea.addCommand("Sheet1!A4:D15", new ImageCommand("image", ImageType.PNG).addArea(imgArea));
                xlsArea.applyAt(new CellRef("Sheet2!A1"), context);
                transformer.write();
                logger.info("written to file");
            }
        }
    }

    public static void execute2() throws IOException {
        try(InputStream is = ImageDemo.class.getResourceAsStream(template2)) {
            try (OutputStream os = new FileOutputStream(output2)) {
                Context context = new Context();
                InputStream imageInputStream = ImageDemo.class.getResourceAsStream("business.png");
                byte[] imageBytes = Util.toByteArray(imageInputStream);
                Department department = new Department("Test Department");
                department.setImage(imageBytes);
                context.putVar("dep", department);
                JxlsHelper.getInstance().processTemplate(is, os, context);
            }
        }
    }

    public static void executeWithNullBytes() throws IOException {
        try(InputStream is = ImageDemo.class.getResourceAsStream(template2)) {
            try (OutputStream os = new FileOutputStream(output3)) {
                Context context = new Context();
                byte[] imageBytes = null;
                Department department = new Department("Test Department");
                department.setImage(imageBytes);
                context.putVar("dep", department);
                JxlsHelper.getInstance().processTemplate(is, os, context);
            }
        }
    }
}
