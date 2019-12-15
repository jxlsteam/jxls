package org.jxls.examples;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.util.List;

import org.junit.Test;
import org.jxls.common.Context;
import org.jxls.entity.Employee;
import org.jxls.util.JxlsHelper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class ObjectCollectionDemo {
    private static final Logger logger = LoggerFactory.getLogger(ObjectCollectionDemo.class);

    @Test
    public void test() throws ParseException, IOException {
        logger.info("Running Object Collection demo");
        List<Employee> employees = Employee.generateSampleEmployeeData();
        try (InputStream is = ObjectCollectionDemo.class.getResourceAsStream("object_collection_template.xls")) {
            try (OutputStream os = new FileOutputStream("target/object_collection_output.xls")) {
                Context context = new Context();
                context.putVar("employees", employees);
                JxlsHelper.getInstance().processTemplate(is, os, context);
            }
        }
    }
}
