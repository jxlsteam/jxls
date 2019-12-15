package org.jxls.examples;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

import org.junit.Test;
import org.jxls.common.Context;
import org.jxls.entity.Department;
import org.jxls.util.JxlsHelper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Created by Leonid Vysochyn on 6/30/2015.
 * todo: improve each command to be able to set merge cells
 */
public class MergedCellsDemo  {
    private static final Logger logger = LoggerFactory.getLogger(MergedCellsDemo.class);
    private static final String template = "merged_cells_demo.xls";
    private static final String output = "target/merged_cells_output.xls";

    @Test
    public void test() throws IOException {
        logger.info("Running merged cells demo");
        List<Department> departments = Department.createDepartments();
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
