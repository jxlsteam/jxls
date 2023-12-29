package org.jxls.unittests;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNull;

import org.junit.Test;
import org.jxls.command.DynamicSheetNameGenerator;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.PoiExceptionThrower;
import org.jxls.logging.JxlsLogger;
import org.jxls.templatebasedtests.multisheet.DynamicSheetNameGeneratorTest;
import org.jxls.transform.SafeSheetNameBuilder;
import org.jxls.transform.poi.PoiSafeSheetNameBuilder;

/**
 * This is a multi sheet test.
 * 
 * @see DynamicSheetNameGeneratorTest
 */
public class DynamicSheetNameGeneratorUnitTest {
    private static final JxlsLogger logger = new PoiExceptionThrower();
    
    /** Old-style test */
    @Test
    public void test() {
        // Prepare
        Context context = new Context(); // No SafeSheetNameBuilder.
        context.putVar("sheetnames", "doe");
        
        // Test
        DynamicSheetNameGenerator gen = new DynamicSheetNameGenerator("sheetnames", new CellRef("A1"));
        
        // Verify
        assertEquals("doe", gen.generateCellRef(0, context, logger).getSheetName());
        assertEquals("'doe(1)'!A1", gen.generateCellRef(1, context, logger).toString());
        assertEquals("'doe(2)'!A1", gen.generateCellRef(2, context, logger).toString());
    }

    @Test
    public void testNull() {
        DynamicSheetNameGenerator gen = new DynamicSheetNameGenerator("N/A", new CellRef("A1"));
        assertNull(gen.generateCellRef(0, new Context(), logger));
    }

    /** New-style test. Setup own sheet name. */
    @Test
    public void testWithSafeNameBuilder() {
        // Prepare
        Context context = new Context();
        context.putVar(SafeSheetNameBuilder.CONTEXT_VAR_NAME, new PoiSafeSheetNameBuilder() {
            @Override
            public String createSafeSheetName(String givenSheetName, int index, JxlsLogger logger) {
                givenSheetName = "#" + (index + 1) + " " + givenSheetName;
                return super.createSafeSheetName(givenSheetName, index, logger);
            }
        });
        context.putVar("sheetnames", "data");
        
        // Test
        DynamicSheetNameGenerator gen = new DynamicSheetNameGenerator("sheetnames", new CellRef("A1"));
        
        // Verify
        assertEquals("#1 data", gen.generateCellRef(0, context, logger).getSheetName());
        assertEquals("#2 data", gen.generateCellRef(1, context, logger).getSheetName());
        assertEquals("#3 data", gen.generateCellRef(2, context, logger).getSheetName());
    }

    /** What does PoiSafeSheetNameBuilder with no modification? */
    @Test
    public void testWithPoiContext() {
        // The PoiSafeSheetNameBuilder acts like the old-style DynamicSheetNameGenerator.
        // Prepare
        Context context = new Context();
        context.putVar(SafeSheetNameBuilder.CONTEXT_VAR_NAME, new PoiSafeSheetNameBuilder());
        context.putVar("sheetnames", "doe");
        
        // Test
        DynamicSheetNameGenerator gen = new DynamicSheetNameGenerator("sheetnames", new CellRef("A1"));
        
        // Verify
        assertEquals("doe", gen.generateCellRef(0, context, logger).getSheetName());
        assertEquals("doe(1)", gen.generateCellRef(1, context, logger).getSheetName());
        assertEquals("doe(2)", gen.generateCellRef(2, context, logger).getSheetName());
    }
}
