package org.jxls.templatebasedtests;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.common.Context;
import org.jxls.entity.Item;

public class Issue173Test {
    
    @Test
    public void test() throws IOException {
        // Prepare
        Context context = new Context();
        List<Item> items = new ArrayList<>();
        items.add(new Item(0, ""));
        context.putVar("items", items);

        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);
        
        // Expectation: no "java.lang.IndexOutOfBoundsException: Index: 1, Size: 1" at StandardFormulaProcessor.java:63
    }
}
