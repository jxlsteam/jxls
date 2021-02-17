package org.jxls.templatebasedtests;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.junit.Assert;
import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;

/**
 * Conditional formatting copying issue
 * 
 * @see IssueB089Test
 * @see ConditionalFormattingTest
 */
public class IssueB110Test {

    @Test
    public void test() throws IOException {
        // Prepare
        List<Item> items = new ArrayList<>();
        items.add(new Item("X", 1, 2));
        items.add(new Item("Y", 3, 4));
        items.add(new Item("Z", 5, 6));
        Context context = new Context();
        context.putVar("items", items);

        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet(0);
            List<String> cfr = w.getConditionalFormattingRanges();
            Assert.assertTrue(cfr.contains("E2"));
            Assert.assertTrue(cfr.contains("E3"));
            Assert.assertFalse(cfr.contains("C2")); // template position
            Assert.assertFalse(cfr.contains("C3")); // template position
        }
    }

    public static class Item {
        public final String label;
        public final int a;
        public final int b;

        public Item(String label, int a, int b) {
            this.label = label;
            this.a = a;
            this.b = b;
        }
    }
}
