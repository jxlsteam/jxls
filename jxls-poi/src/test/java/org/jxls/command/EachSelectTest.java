package org.jxls.command;

import static org.junit.Assert.assertEquals;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.jexl3.JexlBuilder;
import org.apache.commons.jexl3.JexlEngine;
import org.junit.Test;
import org.jxls.common.Context;
import org.jxls.expression.JexlExpressionEvaluator;
import org.jxls.transform.poi.PoiTransformer;
import org.jxls.util.JxlsHelper;

/**
 * issue#157
 */
public class EachSelectTest {

    @Test
    public void test() throws Exception {
        InputStream in = EachSelectTest.class.getResourceAsStream("eachSelect.xlsx");
        File outputFile = new File("target/eachSelect-output.xlsx");
        FileOutputStream out = new FileOutputStream(outputFile);
        Context context = new Context();
        context.putVar("list", getTestData());
        PoiTransformer transformer = PoiTransformer.createTransformer(in, out);
        setupCustomFunctions(transformer);
        JxlsHelper.getInstance().processTemplate(context, transformer);
        
        // Verify
        try (TestWorkbook w = new TestWorkbook(outputFile)) {
            w.selectSheet("eachSelect");
            assertEquals("failed for case: jx:each + jx:if", "A5", w.getCellValueAsString(8, 1)); 
            assertEquals("failed for case: jx:each with select", "A4", w.getCellValueAsString(10, 1)); 
            assertEquals("failed for case: jx:each with select", "A5", w.getCellValueAsString(11, 1)); 
        }
    }

    private List<String> getTestData() {
        List<String> list = new ArrayList<>();
        list.add("A2");
        list.add("A3");
        list.add("A4");
        list.add("A5");
        return list;
    }

    private void setupCustomFunctions(PoiTransformer transformer) {
        Map<String, Object> funcs = new HashMap<>();
        funcs.put("ns", new NSCustomFunctions());
        JexlExpressionEvaluator evaluator = (JexlExpressionEvaluator) transformer.getTransformationConfig().getExpressionEvaluator();
        JexlEngine customJexlEngine = new JexlBuilder().namespaces(funcs).create();
        evaluator.setJexlEngine(customJexlEngine);
    }
    
    public static class NSCustomFunctions {
        
        public Integer func(String str) {
            try {
                return Integer.valueOf(Integer.parseInt(str.substring(1)));
            } catch (NumberFormatException e) {
                return Integer.valueOf(9999);
            }
        }
    }
}
