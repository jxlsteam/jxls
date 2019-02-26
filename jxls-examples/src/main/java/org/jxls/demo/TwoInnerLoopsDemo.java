package org.jxls.demo;

import org.jxls.common.Context;
import org.jxls.demo.guide.Employee;
import org.jxls.demo.model.Department;
import org.jxls.transform.poi.PoiContext;
import org.jxls.util.JxlsHelper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

/**
 * Two nested each commands demo
 * @author Leonid Vysochyn
 */
public class TwoInnerLoopsDemo {
    static Logger logger = LoggerFactory.getLogger(TwoInnerLoopsDemo.class);

    public static void main(String[] args) throws ParseException, IOException {
        logger.info("Running Two Inner Loops demo");
        List<Department> departments = EachIfCommandDemo.createDepartments();
        try(InputStream is = TwoInnerLoopsDemo.class.getResourceAsStream("two_inner_loops_demo.xls")) {
            try (OutputStream os = new FileOutputStream("target/two_inner_loops_demo_output.xls")) {
                Context context = new PoiContext();
                context.putVar("departments", departments);
                JxlsHelper.getInstance().processTemplate(is, os, context);
            }
        }
    }

}
