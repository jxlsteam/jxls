package org.jxls.demo.issue;

import org.jxls.common.Context;
import org.jxls.transform.Transformer;
import org.jxls.util.JxlsHelper;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Arrays;

public class Issue116ExternalFormulasTestCase
{
    private final static String INPUT_FILE_PATH = "IssueExternalFormula_Template.xlsx";

    private final static String OUTPUT_FILE_PATH = "target/IssueExternalFormula_Output.xlsx";

    public static void main(String[] args) throws IOException
    {
        try (InputStream is = Issue116ExternalFormulasTestCase.class.getResourceAsStream(INPUT_FILE_PATH))
        {
            try (OutputStream os = new FileOutputStream(OUTPUT_FILE_PATH))
            {
                Context context = new Context();
                context.putVar("vars", Arrays.asList(1.234, 5.678, 3.1234, 8.9090, 12.34567));
                JxlsHelper jxlsHelper = JxlsHelper.getInstance();
                jxlsHelper.setUseFastFormulaProcessor(true);
                Transformer transformer = jxlsHelper.createTransformer(is, os);
                jxlsHelper.processTemplate(context, transformer);
            }
        }
    }
}