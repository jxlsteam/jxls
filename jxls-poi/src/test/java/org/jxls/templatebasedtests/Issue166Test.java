package org.jxls.templatebasedtests;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.jexl3.JexlBuilder;
import org.apache.commons.jexl3.JexlEngine;
import org.junit.Assert;
import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.JxlsTester.TransformerChecker;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;
import org.jxls.expression.JexlExpressionEvaluator;
import org.jxls.expression.JexlExpressionEvaluatorNoThreadLocal;
import org.jxls.transform.TransformationConfig;
import org.jxls.transform.Transformer;

public class Issue166Test {
    private JexlExpressionEvaluatorNoThreadLocal evaluator;
    
    @Test
    public void createReportTwice_noThreadLocal() {
        createExcelReportUsingCustomFunctions_noThreadLocal();
        createExcelReportUsingCustomFunctions_noThreadLocal();
    }

    @Test
    public void createReportTwice_threadLocal() {
        createExcelReportUsingCustomFunctions_threadLocal();
        createExcelReportUsingCustomFunctions_threadLocal();
    }

    private void createExcelReportUsingCustomFunctions_noThreadLocal() {
        // Prepare
        Context context = createContext();

        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.createTransformerAndProcessTemplate(context, installCustomFunctions_noThreadLocal(context));
        
        // Verify
        verify(tester);

        /* dirty workaround
        try {
            java.lang.reflect.Field field = evaluator.getClass().getDeclaredField("expressionMap");
            field.setAccessible(true);
            Map<?,?> expressionMap = (Map<?,?>) field.get(evaluator);
            expressionMap.clear();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }*/
    }

    private void createExcelReportUsingCustomFunctions_threadLocal() {
        // Prepare
        Context context = createContext();

        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.createTransformerAndProcessTemplate(context, installCustomFunctions_threadLocal(context));
        
        // Verify
        verify(tester);
    }

    public static class MyItem {
        private final int i;
        
        public MyItem(int i) {
            this.i = i;
        }
        
        public String getAbc() {
            return "Hi you " + i + "!";
        }
    }

    private TransformerChecker installCustomFunctions_noThreadLocal(Context context) {
        TransformerChecker tc = new TransformerChecker() {
            @Override
            public Transformer checkTransformer(Transformer transformer) {
                TransformationConfig config = transformer.getTransformationConfig();
                evaluator = new JexlExpressionEvaluatorNoThreadLocal();
                config.setExpressionEvaluator(evaluator);
                Map<String, Object> funcs = new HashMap<>();
                funcs.put("cf", new JXLS2CustomFunctions(context));
                JexlEngine engine = new JexlBuilder().namespaces(funcs).create();
                evaluator.setJexlEngine(engine);
                return transformer;
            }
        };
        return tc;
    }

    private TransformerChecker installCustomFunctions_threadLocal(Context context) {
        TransformerChecker tc = new TransformerChecker() {
            @Override
            public Transformer checkTransformer(Transformer transformer) {
                TransformationConfig config = transformer.getTransformationConfig();
                JexlExpressionEvaluator evaluator_threadLocal = new JexlExpressionEvaluator();
                config.setExpressionEvaluator(evaluator_threadLocal);
                Map<String, Object> funcs = new HashMap<>();
                funcs.put("cf", new JXLS2CustomFunctions(context));
                JexlEngine engine = new JexlBuilder().namespaces(funcs).create();
                evaluator_threadLocal.setJexlEngine(engine);
                return transformer;
            }
        };
        return tc;
    }

    public static class JXLS2CustomFunctions {
        private final Context context;
        
        public JXLS2CustomFunctions(Context context) {
            this.context = context;
        }
        
        public String mach(String varName) {
            return "m_" + getValue(varName);
        }
        
        private Object getValue(String expression) {
            return new JexlExpressionEvaluator(expression).evaluate(context.toMap());
        }
    }

    private Context createContext() {
        Context context = new Context();
        List<MyItem> items = new ArrayList<>();
        for (int i = 1; i <= 30; i++) {
            items.add(new MyItem(i));
        }
        context.putVar("items", items);
        return context;
    }

    private void verify(JxlsTester tester) {
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet(0);
            for (int i = 1; i < 30; i++) {
                Assert.assertEquals("m_Hi you " + i + "!", w.getCellValueAsString(2 + i, 1));
            }
        }
    }
}
