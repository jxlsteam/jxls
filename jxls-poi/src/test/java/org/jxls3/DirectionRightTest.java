package org.jxls3;

import static org.junit.Assert.assertEquals;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.TestWorkbook;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

/**
 * Directin.RIGHT demo
 * <p>Direction is an EachCommand option.
 * Demo shows also use of jx:params(formulaStrategy="BY_COLUMN") for column sums.
 * See also FormulaCopyDemo.</p> 
 */
public class DirectionRightTest {

    @Test
    public void twoColumnsDemo() {
        // Prepare
        List<Transaction> tx = new ArrayList<>();
        tx.add(new Transaction("Apples", new Revenue("West", 1000), new Revenue("South", 900)));
        tx.add(new Transaction("Pears", new Revenue("West", 60), new Revenue("South", 170)));
        tx.add(new Transaction("Peaches", new Revenue("West", 600), new Revenue("South", 2400)));
        tx.add(new Transaction("Plums", new Revenue("West", 500), new Revenue("South", 1400)));
        Map<String, Object> data = new HashMap<>();
        data.put("transactions", tx);
        data.put("regions", List.of(new Region("West", "Paris"), new Region("South", "Malta")));

        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass());
        tester.test(data, JxlsPoiTemplateFillerBuilder.newInstance());
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet("TX");
            assertEquals("West", w.getCellValueAsString(2, 2)); 
            assertEquals("South", w.getCellValueAsString(2, 3)); 
            assertEquals(2160, w.getCellValueAsDouble(8, 2), .005); 
            assertEquals(4870, w.getCellValueAsDouble(8, 3), .005); 
            assertEquals(7030, w.getCellValueAsDouble(8, 4), .005); 
        }
    }

    @Test
    public void fourColumnsDemo() {
        // Prepare
        List<Transaction> tx = new ArrayList<>();
        tx.add(new Transaction("Apples", new Revenue("West", 1000), new Revenue("South", 900), new Revenue("North", 0), new Revenue("East", 0)));
        tx.add(new Transaction("Pears", new Revenue("West", 60), new Revenue("South", 170), new Revenue("North", 500), new Revenue("East", 0)));
        tx.add(new Transaction("Peaches", new Revenue("West", 600), new Revenue("South", 2400), new Revenue("North", 0), new Revenue("East", 36)));
        tx.add(new Transaction("Plums", new Revenue("West", 500), new Revenue("South", 1400), new Revenue("North", 0), new Revenue("East", 0)));
        Map<String, Object> data = new HashMap<>();
        data.put("transactions", tx);
        data.put("regions", List.of(new Region("West", "Paris"), new Region("South", "Malta"), new Region("North", "Oslo"), new Region("East", "Minsk")));

        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass());
        tester.test(data, JxlsPoiTemplateFillerBuilder.newInstance());
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet("TX");
            assertEquals("West", w.getCellValueAsString(2, 2)); 
            assertEquals("South", w.getCellValueAsString(2, 3)); 
            assertEquals("North", w.getCellValueAsString(2, 4)); 
            assertEquals("East", w.getCellValueAsString(2, 5)); 
            assertEquals(2160, w.getCellValueAsDouble(8, 2), .005); 
            assertEquals(4870, w.getCellValueAsDouble(8, 3), .005); 
            assertEquals(500, w.getCellValueAsDouble(8, 4), .005); 
            assertEquals(36, w.getCellValueAsDouble(8, 5), .005); 
            assertEquals(7566, w.getCellValueAsDouble(8, 6), .005); 
        }
    }

    public static class Transaction {
        private String name;
        private final List<Revenue> revenues = new ArrayList<>();
        
        public Transaction(String name, Revenue ...revenues) {
            this.name = name;
            for (Revenue r : revenues) {
                this.revenues.add(r);
            }
        }

        public String getName() {
            return name;
        }
        
        public void setName(String name) {
            this.name = name;
        }
        
        public List<Revenue> getRevenues() {
            return revenues;
        }
    }

    public static class Revenue {
        private String region;
        private double amount;

        public Revenue(String region, double amount) {
            this.region = region;
            this.amount = amount;
        }

        public String getRegion() {
            return region;
        }

        public void setRegion(String region) {
            this.region = region;
        }

        public double getAmount() {
            return amount;
        }

        public void setAmount(double amount) {
            this.amount = amount;
        }
    }
    
    public static class Region {
        private final String name;
        private final String hq;
        
        public Region(String name, String hq) {
            this.name = name;
            this.hq = hq;
        }
        
        public String getName() {
            return name;
        }

        public String getHq() {
            return hq;
        }
    }
}
