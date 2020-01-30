package org.jxls.examples;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Arrays;
import java.util.List;

import org.junit.Test;
import org.jxls.common.Context;
import org.jxls.entity.Department;
import org.jxls.transform.poi.PoiTransformer;
import org.jxls.util.JxlsHelper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * 1st Multi-sheet Markup Demo
 * 
 * @author Leonid Vysochyn
 */
public class MultiSheetMarkup1Demo {
    private static final Logger logger = LoggerFactory.getLogger(MultiSheetMarkup1Demo.class);
    private static final String template = "multisheet_markup_demo.xls";
    private static final String output = "target/multisheet_markup_output.xls";

    @Test
    public void test() throws IOException {
        logger.info("Running Multiple Sheet Markup demo");
        List<Department> departments = Department.createDepartments();
        logger.info("Opening input stream");
        try (InputStream is = MultiSheetMarkup1Demo.class.getResourceAsStream(template)) {
            try (OutputStream os = new FileOutputStream(output)) {
                Context context = PoiTransformer.createInitialContext();
                context.putVar("departments", departments);
                context.putVar("sheetNames", Arrays.asList(
                        departments.get(0).getName(),
                        departments.get(1).getName(),
                        departments.get(2).getName()));
                // with multi sheets it is better to use StandardFormulaProcessor by disabling the FastFormulaProcessor
                JxlsHelper
                        .getInstance()
                        .setUseFastFormulaProcessor(false)
                        .setDeleteTemplateSheet(true)
                        .processTemplate(is, os, context);
            }
        }
    }
}
