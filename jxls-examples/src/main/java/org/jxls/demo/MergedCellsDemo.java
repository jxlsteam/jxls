package org.jxls.demo;

import org.jxls.common.Context;
import org.jxls.demo.model.Department;
import org.jxls.util.JxlsHelper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

/**
 * Created by Leonid Vysochyn on 6/30/2015.
 * todo: improve each command to be able to set merge cells
 */
public class MergedCellsDemo  {
    static Logger logger = LoggerFactory.getLogger(MergedCellsDemo.class);
    private static String template = "merged_cells_demo.xls";
    private static String output = "target/merged_cells_output.xls";

    public static void main(String[] args) throws IOException {
        logger.info("Running merged cells demo");
        execute();
    }

    public static void execute() throws IOException {
        List<Department> departments = EachIfCommandDemo.createDepartments();
        logger.info("Opening input stream");
        try(InputStream is = XlsCommentBuilderDemo.class.getResourceAsStream(template)) {
            try (OutputStream os = new FileOutputStream(output)) {
                Context context = new Context();
                context.putVar("departments", departments);
                JxlsHelper.getInstance().processTemplate(is, os, context);
            }
        }
    }
}
