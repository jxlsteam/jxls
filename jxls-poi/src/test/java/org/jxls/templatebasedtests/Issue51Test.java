package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;

import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;

public class Issue51Test {

    @Test
    public void test() {
        // Prepare
        Context context = new Context();
        
        List<LocalDate> dates = new ArrayList<>();
        dates.add(LocalDate.now().plusDays(-60));
        dates.add(LocalDate.now());
        context.putVar("dates", dates);
        
        List<Equipment> equipmentList = new ArrayList<>();
        Equipment e = new Equipment();
        e.highway = "A40";
        e.add(10); // date -60d
        e.add(11); // date now
        equipmentList.add(e);
        e = new Equipment();
        e.highway = "A57";
        e.add(200); // date -60d
        e.add(1000); // date now
        equipmentList.add(e);
        context.putVar("equipmentList", equipmentList);
        
        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);

        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet(0);
            assertEquals(21d, w.getCellValueAsDouble(7, 9), 0.01d);   // I7
            assertEquals(1200d, w.getCellValueAsDouble(8, 9), 0.01d); // I8
            assertEquals(1221d, w.getCellValueAsDouble(9, 9), 0.01d); // I9
            assertEquals(210d, w.getCellValueAsDouble(9, 7), 0.01d);  // G9 <- problem
            assertEquals(1011d, w.getCellValueAsDouble(9, 8), 0.01d); // H9 <- problem
        }
    }
    
    public static class Equipment {
        public String highway;
        public List<Report> reportList = new ArrayList<>();
        
        public void add(int q) {
            Report r = new Report();
            r.quantity = q;
            reportList.add(r);
        }
    }
    
    public static class Report {
        public int quantity;
    }
}
