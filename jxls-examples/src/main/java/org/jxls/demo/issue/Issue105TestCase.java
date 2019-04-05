package org.jxls.demo.issue;

import org.jxls.common.Context;
import org.jxls.transform.Transformer;
import org.jxls.util.JxlsHelper;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Arrays;

public class Issue105TestCase
{
    private final static String LIST = "vars";

    private final static String INPUT_FILE_PATH = "issue105_template.xlsx";

    private final static String OUTPUT_FILE_PATH = "target/issue105_output.xlsx";

    public static void main(String[] args) throws IOException
    {
        try (InputStream is = Issue105TestCase.class.getResourceAsStream(INPUT_FILE_PATH))
        {
            try (OutputStream os = new FileOutputStream(OUTPUT_FILE_PATH))
            {
                Context context = new Context();
                context.putVar(LIST, Arrays.asList(new Values(), new Values(), new Values()));
                JxlsHelper jxlsHelper = JxlsHelper.getInstance();
                jxlsHelper.setUseFastFormulaProcessor(false);
                Transformer transformer = jxlsHelper.createTransformer(is, os);
                jxlsHelper.processTemplate(context, transformer);
            }
        }
    }

    public static class Values
    {
        public double smallValue = 1.2;

        public double bigValue = 1.3E22;
    }
}