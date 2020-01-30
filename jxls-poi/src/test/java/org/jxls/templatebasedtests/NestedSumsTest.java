package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;
import org.jxls.functions.DoubleSummarizerBuilder;
import org.jxls.functions.GroupSum;

/**
 * Simplified real world testcase for nested sums
 * 
 * @author Marcus Warm (XDEV-3784)
 */
public class NestedSumsTest {
    
    /**
     * Nested group sums.
     * Works with Excel sums.
     */
    @Test
    public void nestedSums() {
        // Prepare
        Context context = new Context();
        List<Map<String, Object>> testData = getTestData();
        testData.get(2).put("class2", "Liegenschaften");
        context.putVar("list", testData);
        
        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass(), "nestedSums");
        tester.processTemplate(context);

        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet("nestedsums");

            assertEquals("Wrong amount in D25!\n", Double.valueOf(123d), w.getCellValueAsDouble(25, 4));

            assertEquals("Wrong group sum in D9!\n", Double.valueOf(1000d), w.getCellValueAsDouble(9, 4));
            assertEquals("Wrong group sum in D15!\n", Double.valueOf(700d), w.getCellValueAsDouble(15, 4));
            assertEquals("Wrong sum in D16!\n", Double.valueOf(1700d), w.getCellValueAsDouble(16, 4));
            
            assertEquals("Wrong group sum in D23!\n", Double.valueOf(600d), w.getCellValueAsDouble(23, 4));
            assertEquals("Wrong group sum in D26!\n", Double.valueOf(123d), w.getCellValueAsDouble(26, 4));
            assertEquals("Wrong sum in D27!\n", Double.valueOf(723d), w.getCellValueAsDouble(27, 4));
            
            assertEquals("Wrong group sum in D32!\n", Double.valueOf(0.31d), w.getCellValueAsDouble(32, 4));
            assertEquals("Wrong sum in D33!\n", Double.valueOf(0.31d), w.getCellValueAsDouble(33, 4));
            
            assertEquals("Wrong grand total! (D35)\n", Double.valueOf(2423.31d), w.getCellValueAsDouble(35, 4));
        }
    }

    /**
     * Nested group sums with jx:if after the mid jx:each to omit 2nd layer.
     * Solution: use GroupSum for 1st layer sum.
     */
    @Test
    public void nestedSums_withIf() {
        // Prepare
        Context context = new Context();
        // We need to calculate the group sum for the part where the children are omitted by the jx:if.
        context.putVar("G", new GroupSum<Double>(context, new DoubleSummarizerBuilder()));
        context.putVar("list", getTestData());
        
        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass(), "nestedSums_withIf");
        tester.processTemplate(context);

        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet("nestedsums");
            assertEquals(Double.valueOf(1700d), w.getCellValueAsDouble(5, 4));
            assertEquals(Double.valueOf( 600d), w.getCellValueAsDouble(12, 4));
            assertEquals(Double.valueOf( 123d), w.getCellValueAsDouble(15, 4));
            assertEquals(Double.valueOf( 723d), w.getCellValueAsDouble(16, 4));
            assertEquals(Double.valueOf(0.31d), w.getCellValueAsDouble(19, 4));
            assertEquals(Double.valueOf(2423.31), w.getCellValueAsDouble(21, 4));
        }
    }

    /**
     * Problem 1:
     * We used 2 jx:each at 1st layer in 1 Excel comment. First group by supertype, second jx:each with group by instrument.
     * The supertype is not displayed in the report.
     * Problem: We have two sums for the same instruments.
     * Solution: This cannot work. We must use only the 2nd jx:each.
     * 
     * Problem 2:
     * Omit class 2 header and footer for bonds.
     * Solution: use GroupSum for 1st and 2nd layer sums.
     */
    @Test
    public void nestedSums_withIf2() {
        // Prepare
        Context context = new Context();
        context.putVar("G", new GroupSum<Double>(context, new DoubleSummarizerBuilder()));
        context.putVar("list", getTestData());

        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass(), "nestedSums_withIf2");
        tester.processTemplate(context);

        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet("nestedsums");
            assertEquals(Double.valueOf(1700d), w.getCellValueAsDouble(12, 4));
            assertEquals(Double.valueOf( 600d), w.getCellValueAsDouble(19, 4));
            assertEquals(Double.valueOf( 123d), w.getCellValueAsDouble(22, 4));
            assertEquals(Double.valueOf( 723d), w.getCellValueAsDouble(23, 4));
            assertEquals(Double.valueOf(2423.31d), w.getCellValueAsDouble(29, 4));
        }
    }
    
    private List<Map<String, Object>> getTestData() {
        List<Map<String, Object>> list = new ArrayList<>();
        add(list, "Commodities type A", "Commodity", "Liegenschaften", "Wolterstr. 100", 250d);
        add(list, "Commodities type A", "Commodity", "Liegenschaften", "Stauffenbergallee", 500d);
        add(list, "Commodities type A", "Commodity", "Immobilien", "Wolterstr. 102", 250d);
        add(list, "Commodities type B", "Commodity", "Fahrzeuge", "Porsche 911", 100d);
        add(list, "Commodities type B", "Commodity", "Fahrzeuge", "Mercedes Maybach", 300d);
        add(list, "Commodities type B", "Commodity", "Fahrzeuge", "Mercedes-Benz SLK 350", 60d);
        add(list, "Commodities type B", "Commodity", "Fahrzeuge", "Bentley Flying Spur", 240d);
        add(list, "Bonds", "Bond", "Base", "AC-100 K1", 200d);
        add(list, "Bonds", "Bond", "Base", "AC-100 K2", 200d);
        add(list, "Bonds", "Bond", "Base", "AC-100 K3", 200d);
        add(list, "Bonds", "Bond", "Super", "MX 12", 123d);
        add(list, "Shares", "Share", "Base", "L77", 0.31d);
        return list;
    }
    
    private void add(List<Map<String, Object>> list, String supertype, String instrument, String class2, String description, double amount) {
        Map<String, Object> map = new HashMap<>();
        map.put("supertype", supertype);
        map.put("instrument", instrument);
        map.put("class2", class2);
        map.put("description", description);
        map.put("amount", Double.valueOf(amount)); // in 1000 EUR
        list.add(map);
    }
}
