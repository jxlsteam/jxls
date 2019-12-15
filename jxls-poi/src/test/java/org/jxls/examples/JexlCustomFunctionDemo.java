package org.jxls.examples;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.jexl3.JexlBuilder;
import org.apache.commons.jexl3.JexlEngine;
import org.junit.Test;
import org.jxls.area.Area;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.expression.JexlExpressionEvaluator;
import org.jxls.transform.Transformer;
import org.jxls.util.TransformerFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * JEXL Custom Function Demo
 * 
 * <p>Alternative: It's also possible to just add a Java object to the context and call a method on this object (using JEXL).</p>
 * 
 * @author Leonid Vysochyn on 22-Jul-15.
 */
public class JexlCustomFunctionDemo {
    private static final Logger logger = LoggerFactory.getLogger(JexlCustomFunctionDemo.class);
    private static final String template = "jexl_custom_function_template.xlsx";
    private static final String output = "target/jexl_custom_function_output.xlsx";

    @Test
    public void test() throws ParseException, IOException {
        logger.info("Running JEXL Custom Function demo");
        try (InputStream is = JexlCustomFunctionDemo.class.getResourceAsStream(template)) {
            try (OutputStream os = new FileOutputStream(output)) {
                Transformer transformer = TransformerFactory.createTransformer(is, os);
                AreaBuilder areaBuilder = new XlsCommentAreaBuilder(transformer);
                List<Area> xlsAreaList = areaBuilder.build();
                Area xlsArea = xlsAreaList.get(0);
                Context context = new Context();
                context.putVar("x", 5);
                context.putVar("y", 10);
                JexlExpressionEvaluator evaluator = (JexlExpressionEvaluator) transformer.getTransformationConfig().getExpressionEvaluator();
                Map<String, Object> functionMap = new HashMap<>();
                functionMap.put("demo", new MyCustomFunctions());
                JexlEngine customJexlEngine = new JexlBuilder().namespaces(functionMap).create();
                evaluator.setJexlEngine(customJexlEngine);
                xlsArea.applyAt(new CellRef("Sheet1!A1"), context);
                transformer.write();
            }
        }
    }

    public static class MyCustomFunctions {
        
        public Integer mySum(Integer x, Integer y) {
            return x + y;
        }
    }
}
