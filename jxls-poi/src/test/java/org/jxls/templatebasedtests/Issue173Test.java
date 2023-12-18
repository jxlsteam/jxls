package org.jxls.templatebasedtests;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.common.Context;
import org.jxls.common.PoiExceptionThrower;
import org.jxls.entity.Item;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

public class Issue173Test {
    
    @Test
    public void test() throws IOException {
        // Prepare
        Context context = new Context();
        List<Item> items = new ArrayList<>();
        items.add(new Item(0, ""));
        context.putVar("items", items);

        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass());
        tester.test(context.toMap(), JxlsPoiTemplateFillerBuilder.newInstance().withLogger(new PoiExceptionThrower() {
            @Override
            public void handleGetObjectPropertyException(Exception e, Object obj, String propertyName) {
                if (!"mun".equals(propertyName) && !"periodEndOfMonth".equals(propertyName)) {
                    super.handleGetObjectPropertyException(e, obj, propertyName);
                }
            }
        }));
        
        // Expectation: no "java.lang.IndexOutOfBoundsException: Index: 1, Size: 1" at StandardFormulaProcessor.java:63
    }
}
