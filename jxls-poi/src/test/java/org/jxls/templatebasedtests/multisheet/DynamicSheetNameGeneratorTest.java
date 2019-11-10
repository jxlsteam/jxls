package org.jxls.templatebasedtests.multisheet;

import java.io.IOException;

import org.junit.Test;
import org.jxls.common.Context;
import org.jxls.unittests.DynamicSheetNameGeneratorUnitTest;

/**
 * This is a multi sheet test.
 * 
 * @see DynamicSheetNameGeneratorUnitTest
 */
public class DynamicSheetNameGeneratorTest extends AbstractMultiSheetTest {

    /**
     * The multisheet attribute has an expression.
     * jx:each(items="sheets", var="sh", multisheet="sh.name", lastCell="A2")
     */
    @Test
    public void testWithExpression() throws IOException {
        Context ctx = new Context();
        ctx.putVar("sheets", getTestSheets());
        createExcel(ctx, "DynamicSheetNameGeneratorTest.xlsx", "target/DynamicSheetNameGeneratorTest_output.xlsx");
        // Result: Sheets with name "data" and "parameters" created.
    }
}
