package org.jxls.demo.issue;

import org.jxls.common.Context;
import org.jxls.demo.guide.Employee;
import org.jxls.demo.guide.ObjectCollectionDemo;
import org.jxls.util.JxlsHelper;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.util.Arrays;
import java.util.Collection;
import java.util.List;

public class Issue127TestCase {

    public static void main(String[] args) throws IOException, ParseException {
        Collection<Integer> datas = Arrays.asList(1,2,3,4);
        try(InputStream is = Issue127TestCase.class.getResourceAsStream("issue127_template.xlsx")) {
            try (OutputStream os = new FileOutputStream("target/issue127_output.xlsx")) {
                Context context = new Context();
                context.putVar("datas", datas);
                JxlsHelper.getInstance().processTemplate(is, os, context);
            }
        }
    }
}
