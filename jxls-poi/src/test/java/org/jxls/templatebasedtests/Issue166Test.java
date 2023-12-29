package org.jxls.templatebasedtests;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.junit.Assert;
import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;
import org.jxls.common.NeedsContext;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

public class Issue166Test {
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
        Map<String, Object> data = new HashMap<>();
        List<MyItem> items = new ArrayList<>();
        for (int i = 1; i <= 30; i++) {
            items.add(new MyItem(i));
        }
        data.put("items", items);
        MyCustomFunctions cf = new MyCustomFunctions();
        data.put("cf", cf);
        
        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass(), "3");
        tester.test(data, JxlsPoiTemplateFillerBuilder.newInstance().withExceptionThrower().needsContext(cf));
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet(0);
            for (int i = 1; i < 30; i++) {
                Assert.assertEquals("m_Hi you " + i + "!", w.getCellValueAsString(2 + i, 1));
            }
        }
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

    public class MyCustomFunctions implements NeedsContext {
        private Context context;
        
        public String mach(String varName) {
            return "m_" + getValue(varName);
        }
        
        private Object getValue(String expression) {
            return context.evaluate(expression);
        }

        @Override
        public void setContext(Context context) {
            this.context = context;
        }
        
        @Override
        public String toString() {
            return "MyCustomFunctions";
        }
    }
}
