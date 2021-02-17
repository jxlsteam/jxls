package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.jexl3.JexlBuilder;
import org.apache.commons.jexl3.JexlEngine;
import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.JxlsTester.TransformerChecker;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;
import org.jxls.expression.JexlExpressionEvaluator;
import org.jxls.transform.Transformer;

/**
 * Issue B157: jx:each select attribute limited expression evaluation context
 */
public class EachSelectTest {

    @Test
    public void test() {
        // Prepare
        Context context = new Context();
        context.putVar("list", getTestData());
        TransformerChecker useTransformerWithCustomFunctions = new TransformerChecker() {
            @Override
            public Transformer checkTransformer(Transformer transformer) {
                setupCustomFunctions(transformer);
                return transformer;
            }

            private void setupCustomFunctions(Transformer transformer) {
                Map<String, Object> funcs = new HashMap<>();
                funcs.put("ns", new NSCustomFunctions());
                JexlExpressionEvaluator evaluator = (JexlExpressionEvaluator) transformer.getTransformationConfig().getExpressionEvaluator();
                JexlEngine customJexlEngine = new JexlBuilder().namespaces(funcs).create();
                evaluator.setJexlEngine(customJexlEngine);
            }
        };

        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.createTransformerAndProcessTemplate(context, useTransformerWithCustomFunctions);

        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
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
