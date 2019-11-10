package org.jxls.unittests;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNull;

import org.junit.Test;
import org.jxls.command.DynamicSheetNameGenerator;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.templatebasedtests.multisheet.AbstractMultiSheetTest.TestExpressionEvaluator;
import org.jxls.templatebasedtests.multisheet.DynamicSheetNameGeneratorTest;
import org.jxls.transform.SafeSheetNameBuilder;
import org.jxls.transform.poi.PoiContext;
import org.jxls.transform.poi.PoiSafeSheetNameBuilder;

/**
 * This is a multi sheet test.
 * 
 * @see DynamicSheetNameGeneratorTest
 */
public class DynamicSheetNameGeneratorUnitTest {

    /** Old-style test */
    @Test
    public void test() {
        // Prepare
        Context context = new Context(); // No SafeSheetNameBuilder.
        context.putVar("sheetnames", "doe");
        
        // Test
        DynamicSheetNameGenerator gen = new DynamicSheetNameGenerator("sheetnames", new CellRef("A1"), new TestExpressionEvaluator());
        
        // Verify
        assertEquals("doe", gen.generateCellRef(0, context).getSheetName());
        assertEquals("'doe(1)'!A1", gen.generateCellRef(1, context).toString());
        assertEquals("'doe(2)'!A1", gen.generateCellRef(2, context).toString());
    }

    @Test
    public void testNull() {
        DynamicSheetNameGenerator gen = new DynamicSheetNameGenerator("N/A", new CellRef("A1"), new TestExpressionEvaluator());
        assertNull(gen.generateCellRef(0, new Context()));
    }

    /** New-style test. Setup own sheet name. */
    @Test
    public void testWithSafeNameBuilder() {
        // Prepare
        Context context = new Context();
        context.putVar(SafeSheetNameBuilder.CONTEXT_VAR_NAME, new PoiSafeSheetNameBuilder() {
            @Override
            public String createSafeSheetName(String givenSheetName, int index) {
                givenSheetName = "#" + (index + 1) + " " + givenSheetName;
                return super.createSafeSheetName(givenSheetName, index);
            }
        });
        context.putVar("sheetnames", "data");
        
        // Test
        DynamicSheetNameGenerator gen = new DynamicSheetNameGenerator("sheetnames", new CellRef("A1"), new TestExpressionEvaluator());
        
        // Verify
        assertEquals("#1 data", gen.generateCellRef(0, context).getSheetName());
        assertEquals("#2 data", gen.generateCellRef(1, context).getSheetName());
        assertEquals("#3 data", gen.generateCellRef(2, context).getSheetName());
    }

    /** What does PoiSafeSheetNameBuilder with no modification? */
    @Test
    public void testWithPoiContext() {
        // Prepare
        Context context = new PoiContext(); // The PoiSafeSheetNameBuilder acts like the old-style DynamicSheetNameGenerator.
        context.putVar("sheetnames", "doe");
        
        // Test
        DynamicSheetNameGenerator gen = new DynamicSheetNameGenerator("sheetnames", new CellRef("A1"), new TestExpressionEvaluator());
        
        // Verify
        assertEquals("doe", gen.generateCellRef(0, context).getSheetName());
        assertEquals("doe(1)", gen.generateCellRef(1, context).getSheetName());
        assertEquals("doe(2)", gen.generateCellRef(2, context).getSheetName());
    }
}
