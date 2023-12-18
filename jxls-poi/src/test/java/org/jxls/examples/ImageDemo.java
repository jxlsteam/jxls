package org.jxls.examples;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.junit.Ignore;
import org.junit.Test;
import org.jxls.area.XlsArea;
import org.jxls.command.ImageCommand;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.ImageType;
import org.jxls.entity.Department;
import org.jxls.transform.Transformer;
import org.jxls.util.JxlsHelper;
import org.jxls.util.TransformerFactory;

/**
 * @author Leonid Vysochyn
 */
public class ImageDemo {
    private static final String template = "image_demo.xlsx";
    private static final String template2 = "image_demo2.xlsx";
    private static final String output = "target/image_output.xlsx";
    private static final String output2 = "target/image_output2.xlsx";
    private static final String output3 = "target/image_output3.xlsx";
    private static final String template4 = "image_demo4.xlsx";
    private static final String output4 = "target/image_output4.xlsx";

    @Ignore // TODO later because demo
    @Test
    public void test() throws IOException {
        execute();
        execute2();
        executeWithNullBytes();
        executeResizePicture();
    }

    private void execute() throws IOException {
        try (InputStream is = ImageDemo.class.getResourceAsStream(template)) {
            try (OutputStream os = new FileOutputStream(output)) {
                Transformer transformer = TransformerFactory.createTransformer(is, os);
                XlsArea xlsArea = new XlsArea("Sheet1!A1:N30", transformer);
                Context context = new Context();
                InputStream imageInputStream = ImageDemo.class.getResourceAsStream("business.png");
                byte[] imageBytes = toByteArray(imageInputStream);
                context.putVar("image", imageBytes);
                XlsArea imgArea = new XlsArea("Sheet1!A5:D15", transformer);
                xlsArea.addCommand("Sheet1!A4:D15", new ImageCommand("image", ImageType.PNG).addArea(imgArea));
                xlsArea.applyAt(new CellRef("Sheet2!A1"), context);
                transformer.write();
            }
        }
    }

    private void execute2() throws IOException {
        try (InputStream is = ImageDemo.class.getResourceAsStream(template2)) {
            try (OutputStream os = new FileOutputStream(output2)) {
                Context context = new Context();
                InputStream imageInputStream = ImageDemo.class.getResourceAsStream("business.png");
                byte[] imageBytes = toByteArray(imageInputStream);
                Department department = new Department("Test Department");
                department.setImage(imageBytes);
                context.putVar("dep", department);
                JxlsHelper.getInstance().processTemplate(is, os, context);
            }
        }
    }

    private void executeWithNullBytes() throws IOException {
        try (InputStream is = ImageDemo.class.getResourceAsStream(template2)) {
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

    private void executeResizePicture() throws IOException {
        try (InputStream is = ImageDemo.class.getResourceAsStream(template4)) {
            try (OutputStream os = new FileOutputStream(output4)) {
                Context context = new Context();
                InputStream imageInputStream = ImageDemo.class.getResourceAsStream("business.png");
                byte[] imageBytes = toByteArray(imageInputStream);
                Department department = new Department("Test Department");
                department.setImage(imageBytes);
                context.putVar("dep", department);
                JxlsHelper.getInstance().processTemplate(is, os, context);
            }
        }
    }
    
    /**
     * Reads all the data from the input stream, and returns the bytes read.
     * 
     * @param stream -
     * @return byte array
     * @throws IOException -
     */
    private byte[] toByteArray(InputStream stream) throws IOException {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        byte[] buffer = new byte[4096];
        int read = 0;
        while (read != -1) {
            read = stream.read(buffer);
            if (read > 0) {
                baos.write(buffer, 0, read);
            }
        }
        return baos.toByteArray();
    }
}
