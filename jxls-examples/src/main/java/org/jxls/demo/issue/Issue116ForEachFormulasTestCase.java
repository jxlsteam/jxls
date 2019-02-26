package org.jxls.demo.issue;

import org.jxls.common.Context;
import org.jxls.expression.JexlExpressionEvaluator;
import org.jxls.transform.Transformer;
import org.jxls.util.JxlsHelper;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

public class Issue116ForEachFormulasTestCase
{
    private final static String INPUT_FILE_PATH = "IssueForEachFormula_Template.xlsx";

    private final static String OUTPUT_FILE_PATH = "target/IssueForEachFormula_Output.xlsx";

    public static void main(String[] args) throws IOException
    {
        try (InputStream is = Issue116ForEachFormulasTestCase.class.getResourceAsStream(INPUT_FILE_PATH))
        {
            try (OutputStream os = new FileOutputStream(OUTPUT_FILE_PATH))
            {
                Context context = new Context();
                context.putVar("vars", Arrays.asList(1.234, 5.678, 3.1234, 8.9090, 12.34567));
                JxlsHelper jxlsHelper = JxlsHelper.getInstance();
                jxlsHelper.setUseFastFormulaProcessor(false);
                Transformer transformer = jxlsHelper.createTransformer(is, os);
                final JexlExpressionEvaluator evaluator = (JexlExpressionEvaluator) transformer.getTransformationConfig().getExpressionEvaluator();
                Map<String, Object> jxlsMap = new HashMap<>();
                jxlsMap.put("math", Math.class); // add Math utility functions

//                evaluator.getJexlEngine().setFunctions(jxlsMap);
                jxlsHelper.processTemplate(context, transformer);
            }
        }
    }
}