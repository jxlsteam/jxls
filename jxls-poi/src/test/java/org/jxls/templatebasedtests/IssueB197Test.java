package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import java.util.Arrays;
import java.util.Collections;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.common.CellData;
import org.jxls.common.CellData.CellType;
import org.jxls.common.CellRef;
import org.jxls.common.Context;

/**
 * Jointed cell references do not work anymore with several empty collections.
 * Bug in 2.7.0 caused by commit bfdf618295cfa8ae68e72f67a428e69ec11c3e72.
 * 
 * @author Christian Raack
 */
public class IssueB197Test {

    @Test
    public void test() {
        // Prepare
        Context context = new Context();
        context.putVar("hansCar", Collections.EMPTY_LIST);
        context.putVar("hansFood", Collections.EMPTY_LIST);
        context.putVar("hansSchool", Arrays.asList(new CostObject("Public Transport", 450d), new CostObject("Books", 300d), new CostObject("Teacher", 3500d)));
        // Empty lists in the Gregor part will cause a wrong sum in template cell B15 with formula $[SUM(U_(B12,B13,B14))]
        context.putVar("gregorCar", Collections.EMPTY_LIST);
        context.putVar("gregorFood", Collections.EMPTY_LIST);
        context.putVar("gregorApparel", Collections.EMPTY_LIST);

        // Test
        JxlsTester tester = JxlsTester.xls(getClass());
        tester.processTemplate(context);
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet(0);
            assertEquals(0d, w.getCellValueAsDouble(12, 2), 0.01d);
            assertEquals(4250d, w.getCellValueAsDouble(7, 2), 0.01d);
            assertEquals("Teacher", w.getCellValueAsString(6, 1));
        }
    }

    public static class CostObject {
        private final String name;
        private final double cost;

        CostObject(String name, double cost) {
            this.name = name;
            this.cost = cost;
        }

        public String getName() {
            return name;
        }

        public double getCost() {
            return cost;
        }
    }
    
    @Test
    public void testU_() {
        CellData cellData = new CellData(new CellRef(1, 1), CellType.FORMULA, "$[CHATEAU_CHAMBRES * 0.9]");
        assertTrue(cellData.isParameterizedFormulaCell());
        assertFalse(cellData.isJointedFormulaCell());
    }
}
