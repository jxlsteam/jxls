package org.jxls.demo.guide;

import org.jxls.common.Context;
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
 * Object collection output demo
 * @author Leonid Vysochyn
 */
public class ObjectCollectionFormulasDemo {
    static Logger logger = LoggerFactory.getLogger(ObjectCollectionFormulasDemo.class);

    public static void main(String[] args) throws ParseException, IOException {
        logger.info("Running Object Collection Formulas demo");
        List<Employee> employees = ObjectCollectionDemo.generateSampleEmployeeData();
        try(InputStream is = ObjectCollectionFormulasDemo.class.getResourceAsStream("formulas_template.xls")) {
            try (OutputStream os = new FileOutputStream("target/formulas_output.xls")) {
                Context context = new Context();
                context.putVar("employees", employees);
                JxlsHelper.getInstance().processTemplateAtCell(is, os, context, "Result!A1");
            }
        }
    }

}
