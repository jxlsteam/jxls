package org.jxls.demo.issue;

import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

public class Issue106TestCase {

    public static void main(String[] args) throws IOException {

        try(InputStream is = Issue106TestCase.class.getResourceAsStream("issue106_template.xls")) {
            try (OutputStream os = new FileOutputStream("target/issue106_output.xls")) {
                Context context = new Context();
                JxlsHelper.getInstance().processTemplate(is, os, context);
            }
        }
    }
}
