package org.jxls.examples;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;

import org.junit.Test;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class IfCommandDemo {
    private static final Logger logger = LoggerFactory.getLogger(IfCommandDemo.class);

    @Test
    public void test() throws ParseException, IOException {
        logger.info("Running IfCommand demo");
        try (InputStream is = ObjectCollectionDemo.class.getResourceAsStream("if_template.xlsx")) {
            try (OutputStream os = new FileOutputStream("target/if_output.xlsx")) {
                Context context = new Context();
                JxlsHelper.getInstance().processTemplate(is, os, context);
            }
        }
    }
}
