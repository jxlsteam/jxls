package org.jxls.demo;

import org.jxls.common.Context;
import org.jxls.demo.guide.Employee;
import org.jxls.demo.model.Department;
import org.jxls.demo.model.Org;
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
 * Formula copy demo
 * @author Leonid Vysochyn
 */
public class FormulaCopyDemo {
    static Logger logger = LoggerFactory.getLogger(FormulaCopyDemo.class);

    public static void main(String[] args) throws ParseException, IOException {
        logger.info("Running Formula Copy demo");
        List<Org> orgs = Org.generate(3, 3);
        try(InputStream is = FormulaCopyDemo.class.getResourceAsStream("formula_copy_template.xls")) {
            try (OutputStream os = new FileOutputStream("target/formula_copy_output.xls")) {
                Context context = new Context();
                context.putVar("orgs", orgs);
                JxlsHelper jxlsHelper = JxlsHelper.getInstance();
                jxlsHelper.setUseFastFormulaProcessor(false);
                jxlsHelper.processTemplate(is, os, context);
            }
        }
    }

}
