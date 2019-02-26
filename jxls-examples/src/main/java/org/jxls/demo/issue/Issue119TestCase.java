package org.jxls.demo.issue;

import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

public class Issue119TestCase {
    public static void main(String[] args) throws IOException {


        try(InputStream is = Issue122TestCase.class.getResourceAsStream("issue119_template.xls")) {
            try (OutputStream os = new FileOutputStream("target/issue119_output.xls")) {
                Context context = new Context();
                context.putVar("title", "Report XLS");
                context.putVar("value", new Integer(100));
                JxlsHelper.getInstance().processTemplate(is, os, context);
            }
        }
    }

}
