package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;

/**
 * Tests if/else - each - if/else combination with an Excel file using Excel comments.
 * 
 * <p>The Excel template contains an English and a German version. Inside the list a row can be a sell or a buy row.</p>
 */
public class If01Test {

    /** Tests the English version with many rows of type buy and sell. */
    @Test
    public void testEnglish() {
        // Prepare
        Context context = new Context();
        context.putVar("lang", "en");
        context.putVar("list", getTestData());

        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);

        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet("if01");
            assertEquals("English report", w.getCellValueAsString(1, 1));
            assertEquals("Buy", w.getCellValueAsString(4, 5));
            assertEquals("Sell", w.getCellValueAsString(5, 5));
            assertEquals("Buy", w.getCellValueAsString(6, 5));
            assertEquals("Subject", w.getCellValueAsString(3, 1));
            assertEquals(123.45d, w.getCellValueAsDouble(9, 2), 0.005d);
            assertEquals(678d, w.getCellValueAsDouble(9, 3), 0.005d);
            assertEquals("Cell must be empty", 0d, w.getCellValueAsDouble(11, 1), 0.005d);
        }
    }

    /** Tests the German version with many rows of type buy and sell. */
    @Test
    public void testGerman() {
        Context context = new Context();
        context.putVar("lang", "de");
        context.putVar("list", getTestData());

        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);

        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet("if01");
            assertEquals("Deutscher Bericht", w.getCellValueAsString(1, 1));
            assertEquals("Kauf", w.getCellValueAsString(4, 5));
            assertEquals("Verkauf", w.getCellValueAsString(5, 5));
            assertEquals("Kauf", w.getCellValueAsString(6, 5));
            assertEquals("Ware", w.getCellValueAsString(3, 1));
            assertEquals(123.45d, w.getCellValueAsDouble(9, 2), 0.005d);
            assertEquals(678d, w.getCellValueAsDouble(9, 3), 0.005d);
            assertEquals("Cell must be empty", 0d, w.getCellValueAsDouble(11, 1), 0.005d);
        }
    }

    /** Tests empty list */
    @Test
    public void testEmpty() {
        Context context = new Context();
        context.putVar("lang", "en");
        context.putVar("list", new ArrayList<Commodity>());

        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);

        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet("if01");
            assertEquals("English report", w.getCellValueAsString(1, 1));
            assertEquals("Subject", w.getCellValueAsString(3, 1));
            assertEquals(0d, w.getCellValueAsDouble(4, 2), 0.005d);
            assertEquals("Cell must be empty", 0d, w.getCellValueAsDouble(11, 1), 0.005d);
        }
    }

    /** Tests list with 1 row. */
    @Test
    public void test1row() {
        Context context = new Context();
        context.putVar("lang", "en");
        List<Commodity> testData = new ArrayList<Commodity>();
        testData.add(new Commodity("1 row", 10d, 100d, "buy"));
        context.putVar("list", testData);

        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);

        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet("if01");
            assertEquals(10d, w.getCellValueAsDouble(4, 2), 0.005d);
            assertEquals("Buy", w.getCellValueAsString(4, 5));
            assertEquals("Cell must be empty", 0d, w.getCellValueAsDouble(11, 1), 0.005d);
        }
    }

    private List<Commodity> getTestData() {
        List<Commodity> ret = new ArrayList<>();
        ret.add(new Commodity("Gas", 1d, 1d, "buy"));
        ret.add(new Commodity("Oil", 2.1d, 10d, "sell"));
        ret.add(new Commodity("Gas", 3.12d, 100d, "buy"));
        ret.add(new Commodity("Gas", 10d, 1000d, "buy"));
        ret.add(new Commodity("Gas", 10d, 1234d, "buy"));
        ret.add(new Commodity("Gas 123", 123.45d, 678d, "sell"));
        return ret;
    }

    static class Commodity extends HashMap<String, Object> {
        private static final long serialVersionUID = -2710058579037153177L;

        public Commodity(String subject, double price, double weight, String sellBuy) {
            put("subject", subject);
            put("price", price);
            put("weight", weight);
            put("sellBuy", sellBuy);
        }
    }
}
