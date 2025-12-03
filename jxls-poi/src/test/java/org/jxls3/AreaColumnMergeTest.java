package org.jxls3;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.junit.Assert;
import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.TestWorkbook;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

public class AreaColumnMergeTest {
    
    @Test
    public void test() {
        // Prepare
        Map<String, Object> data = new HashMap<>();
        data.put("groups", TestDataGenerator.generateTestData()); // groups and transactions
        
        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass());
        tester.test(data, JxlsPoiTemplateFillerBuilder.newInstance());
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet("Transactions");
            Assert.assertEquals("A1:H1,A3:A5,A6:A7,G3:G5,G6:G7,H3:H5,H6:H7", w.getMergedCells());
        }
    }
    
    public static class Transaction {
        private final String group;
        private final String description;
        private final double amount1;
        private final double taxRate1;
        private final double amount2;
        private final double taxRate2;

        public Transaction(String group, String description, double amount1, double taxRate1, double amount2,
                double taxRate2) {
            this.group = group;
            this.description = description;
            this.amount1 = amount1;
            this.taxRate1 = taxRate1;
            this.amount2 = amount2;
            this.taxRate2 = taxRate2;
        }

        public double calculateTotalNetAmount() {
            return amount1 + amount2;
        }

        public double calculateTotalGrossAmount() {
            double gross1 = amount1 * (1 + taxRate1 / 100.0);
            double gross2 = amount2 * (1 + taxRate2 / 100.0);
            return gross1 + gross2;
        }

        public String getGroup() {
            return group;
        }

        public String getDescription() {
            return description;
        }

        public double getAmount1() {
            return amount1;
        }

        public double getTaxRate1() {
            return taxRate1;
        }

        public double getAmount2() {
            return amount2;
        }

        public double getTaxRate2() {
            return taxRate2;
        }
    }
    
    public static class Group {
        private final String group;
        private final List<Transaction> transactions = new ArrayList<>();
        private double totalNet = 0.0;
        private double totalGross = 0.0;

        public Group(String group) {
            this.group = group;
        }

        public void addTransaction(Transaction transaction) {
            this.transactions.add(transaction);
            this.totalNet += transaction.calculateTotalNetAmount();
            this.totalGross += transaction.calculateTotalGrossAmount();
        }

        public String getGroup() {
            return group;
        }

        public double getNet() {
            return totalNet;
        }

        public double getGross() {
            return totalGross;
        }

        public List<Transaction> getTransactions() {
            return transactions;
        }
    }
    
    public static class TestDataGenerator {
        private static final String GROUP_SCOTCH = "Scotch whisky";
        private static final String GROUP_AMERICAN = "American whiskey";
        private static final double TAX_HIGH = 19.0;
        private static final double TAX_MED = 10.0;
        private static final double TAX_LOW = 7.0;

        public static List<Group> generateTestData() {
            Group groupA = new Group(GROUP_SCOTCH);
            Group groupB = new Group(GROUP_AMERICAN);
            
            groupA.addTransaction(new Transaction(
                GROUP_SCOTCH, 
                "Cragganmore 12, 4+1 bottles",
                150.00, TAX_HIGH,
                32.99, TAX_LOW
            ));
            
            groupA.addTransaction(new Transaction(
                GROUP_SCOTCH, 
                "Lagavulin 16, 7 bottles, 1 D.E.",
                524.30, TAX_HIGH,
                92.90, TAX_MED
            ));
            
            groupA.addTransaction(new Transaction(
                GROUP_SCOTCH, 
                "Auchentoshan Three Wood",
                42.90, TAX_HIGH,
                0.00, TAX_LOW
            ));

            groupB.addTransaction(new Transaction(
                GROUP_AMERICAN,
                "Makers Mark",
                120.00, TAX_HIGH,
                35.50, TAX_LOW
            ));
            
            groupB.addTransaction(new Transaction(
                GROUP_AMERICAN,
                "Wild Turkey 101",
                75.00, TAX_LOW,
                10.00, TAX_MED
            ));
            return List.of(groupA, groupB);
        }
    }
}
