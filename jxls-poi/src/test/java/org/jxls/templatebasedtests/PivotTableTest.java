package org.jxls.templatebasedtests;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.junit.Rule;
import org.junit.Test;
import org.jxls.EnglishTestRule;
import org.jxls.JxlsTester;
import org.jxls.JxlsTester.TransformerChecker;
import org.jxls.common.Context;
import org.jxls.transform.Transformer;
import org.jxls.transformer.TestPTTransformer;

public class PivotTableTest {
    @Rule
    public EnglishTestRule english = new EnglishTestRule();

    /**
     * Issue 155: Pivot table does not work with NLS
     */
    @Test
    public void nls() {
        // Prepare
        final Context context = new Context();
        context.putVar("R", getResources()); // NLS
        context.putVar("list", getTestData());
        TransformerChecker useMyTransformer = new TransformerChecker() {
            @Override
            public Transformer checkTransformer(Transformer transformer) {
                return new TestPTTransformer(transformer, context);
            }
        };
        
        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.createTransformerAndProcessTemplate(context, useMyTransformer);
        
        // Verify
        // result: broken PivotTable
        // TODO
    }

    private Map<String, String> getResources() {
        Map<String, String> r = new HashMap<>();
        r.put("name", "Name (EN)");
        r.put("city", "City (EN)");
        return r;
    }

    private List<Map<String, String>> getTestData() {
        List<Map<String, String>> list = new ArrayList<>();
        add(list, "Leonid", "Danzig");
        add(list, "Heil", "Berlin");
        add(list, "Marcus", "Krefeld");
        add(list, "Merkel", "Berlin");
        add(list, "Seehofer", "Berlin");
        add(list, "Waldemar", "Krefeld");
        return list;
    }
    
    private void add(List<Map<String, String>> list, String name, String city) {
        Map<String, String> map = new HashMap<>();
        map.put("name", name);
        map.put("city", city);
        list.add(map);
    }
}
