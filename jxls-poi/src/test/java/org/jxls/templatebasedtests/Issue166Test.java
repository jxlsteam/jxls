package org.jxls.templatebasedtests;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.junit.Assert;
import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;
import org.jxls.expression.JexlExpressionEvaluator;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

public class Issue166Test {
    private Map<String, Object> data;
    // Removed namespace based testcases because it's not recommended to use custom functions that way. createExcelReport() shows the easier way.

    /**
     * testcase with thread local and without "namespace" custom functions.
     * We use custom functions here as POJO object in the context. (syntax: cf.mach)
     */
    @Test
    public void createReportTwice() {
        createExcelReport();
        createExcelReport();
    }

    private void createExcelReport() {
        // Prepare
        Context context = createContext();
        context.putVar("cf", new MyCustomFunctions());
        data = context.toMap();
        
        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass(), "3");
        tester.test(data, JxlsPoiTemplateFillerBuilder.newInstance().withExceptionThrower());
        
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
        
        @Override
        public String toString() {
            return "MyItem:" + getAbc();
        }
    }

    public class MyCustomFunctions {
        
        public String mach(String varName) {
            return "m_" + getValue(varName);
        }
        
        private Object getValue(String expression) {
            if (!data.containsKey("e")) {
                // XXX Geht derzeit nicht, wegen data Zugriffsproblem.
                System.err.println("Running var 'e' does not exist in data!  " + data.keySet());
            }
            return new JexlExpressionEvaluator(expression).evaluate(data);
        }
        
        @Override
        public String toString() {
            return "JXLS2CustomFunctions";
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

    private void verify(Jxls3Tester tester) {
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet(0);
            for (int i = 1; i < 30; i++) {
                Assert.assertEquals("m_Hi you " + i + "!", w.getCellValueAsString(2 + i, 1));
            }
        }
    }
}
