package org.jxls.examples;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.util.List;

import org.junit.Test;
import org.jxls.common.Context;
import org.jxls.entity.Department;
import org.jxls.transform.poi.PoiContext;
import org.jxls.util.JxlsHelper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Two nested each commands demo
 * 
 * @author Leonid Vysochyn
 */
public class TwoInnerLoopsDemo {
    private static final Logger logger = LoggerFactory.getLogger(TwoInnerLoopsDemo.class);

    @Test
    public void test() throws ParseException, IOException {
        logger.info("Running Two Inner Loops demo");
        List<Department> departments = Department.createDepartments();
        try (InputStream is = TwoInnerLoopsDemo.class.getResourceAsStream("two_inner_loops_demo.xls")) {
            try (OutputStream os = new FileOutputStream("target/two_inner_loops_demo_output.xls")) {
                Context context = new PoiContext();
                context.putVar("departments", departments);
                JxlsHelper.getInstance().processTemplate(is, os, context);
            }
        }
    }
}
