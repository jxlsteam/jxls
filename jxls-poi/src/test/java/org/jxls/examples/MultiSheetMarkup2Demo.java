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

/**
 * 2nd Multi-sheet Markup Demo
 * 
 * @author Leonid Vysochyn
 */
public class MultiSheetMarkup2Demo {
    private static final String template = "multisheet_markup_demo-2.xlsx";
    private static final String output = "target/multisheet_markup_output-2.xlsx";

    @Test
    public void test() throws IOException {
        List<Department> departments = Department.createDepartments();
        try (InputStream is = MultiSheetMarkup2Demo.class.getResourceAsStream(template)) {
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
