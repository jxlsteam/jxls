package org.jxls.templatebasedtests;

import java.util.ArrayList;

import org.junit.Assert;
import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;

/**
 * Problem with sums out of empty lists and multiple sheets
 * 
 * @see IssueB166Test
 * @see IssueB188Test
 */
public class IssueB210Test {
    // to run the test with non-empty lists set the flag to true
    // AND make sure the formulas in the template file correspond to your Excel locale
    private static final boolean WITH_DATA = false;
    
    @Test
    public void standard() throws Exception {
        check(false);
    }
    
    @Test
    public void fast() throws Exception {
        check(true);
    }

    public void check(boolean useFastFormulaProcessor) throws Exception {
        // Prepare
        Context context = new Context();
        ArrayList<Item> list = new ArrayList<>();
        if (WITH_DATA) list.add(new Item());
        context.putVar("list1", list);
        list = new ArrayList<>();
        if (WITH_DATA) list.add(new Item());
        context.putVar("list2", list);
        list = new ArrayList<>();
        if (WITH_DATA) list.add(new Item());
        context.putVar("list3", list);
        
        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.setUseFastFormulaProcessor(useFastFormulaProcessor);
        tester.processTemplate(context);
        
        // Verify
        if (WITH_DATA) return;
        try (TestWorkbook w = tester.getWorkbook()) {
            for (int row = 5; row <= 6; row++) {
                w.selectSheet("Sheet1");
                Assert.assertEquals("Error in Sheet1!B" + row, 0d, w.getCellValueAsDouble(row, 2), 0.005d); 
                Assert.assertEquals("Error in Sheet1!C" + row, 0d, w.getCellValueAsDouble(row, 3), 0.005d);
                w.selectSheet("Sheet2");
                Assert.assertEquals("Error in Sheet2!B" + row, 0d, w.getCellValueAsDouble(row, 2), 0.005d); 
                Assert.assertEquals("Error in Sheet2!C" + row, 0d, w.getCellValueAsDouble(row, 3), 0.005d); 
            }
        }
    }
    
    public static class Item {
        private final double value = 1;
        private final double value2 = 2;

        public double getValue() {
            return value;
        }

        public double getValue2() {
            return value2;
        }
    }
}
