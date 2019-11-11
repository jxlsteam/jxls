package org.jxls.templatebasedtests;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.common.Context;

/**
 * Conditional formatting copying issue
 */
public class Issue110Test {

    @Test
    public void test() throws IOException {
        List<Item> items = new ArrayList<>();
        items.add(new Item("X", 1, 2));
        items.add(new Item("Y", 3, 4));
        items.add(new Item("Z", 5, 6));
        Context context = new Context();
        context.putVar("items", items);

        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);
        
        // TODO assertions
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
