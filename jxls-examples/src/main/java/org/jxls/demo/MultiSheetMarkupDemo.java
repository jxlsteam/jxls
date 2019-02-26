package org.jxls.demo;

import org.jxls.common.Context;
import org.jxls.demo.model.Department;
import org.jxls.transform.poi.PoiTransformer;
import org.jxls.util.JxlsHelper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Arrays;
import java.util.List;

/**
 * @author Leonid Vysochyn
 */
public class MultiSheetMarkupDemo {
    static Logger logger = LoggerFactory.getLogger(MultiSheetMarkupDemo.class);
    private static String template = "multisheet_markup_demo.xls";
    private static String output = "target/multisheet_markup_output.xls";

    public static void main(String[] args) throws IOException {
        logger.info("Running Multiple Sheet Markup demo");
        execute();
    }

    public static void execute() throws IOException {
        List<Department> departments = EachIfCommandDemo.createDepartments();
        logger.info("Opening input stream");
        try (InputStream is = MultiSheetMarkupDemo.class.getResourceAsStream(template)) {
            try (OutputStream os = new FileOutputStream(output)) {
                Context context = PoiTransformer.createInitialContext();
                context.putVar("departments", departments);
                context.putVar("sheetNames", Arrays.asList(
                        departments.get(0).getName(),
                        departments.get(1).getName(),
                        departments.get(2).getName()));
                // with multi sheets it is better to use StandardFormulaProcessor by disabling the FastFormulaProcessor
                JxlsHelper.getInstance().setUseFastFormulaProcessor(false).processTemplate(is, os, context);
            }
        }
    }

}
