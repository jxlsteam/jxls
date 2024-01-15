package org.jxls.templatebasedtests.multisheet;

import org.junit.Assert;
import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;
import org.jxls.common.ContextImpl;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;
import org.jxls.unittests.DynamicSheetNameGeneratorUnitTest;
import org.jxls3.MultiSheetTest;

/**
 * This is a multi sheet test.
 * 
 * @see DynamicSheetNameGeneratorUnitTest
 * @see MultiSheetTest
 */
public class DynamicSheetNameGeneratorTest extends AbstractMultiSheetTest {

    /**
     * The multisheet attribute has an expression.
     * jx:each(items="sheets", var="sh", multisheet="sh.name", lastCell="A2")
     */
    @Test
    public void testWithExpression() {
        // Prepare
        Context context = new ContextImpl();
        context.putVar("sheets", getTestSheets());
        
        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass());
        tester.test(context.toMap(), JxlsPoiTemplateFillerBuilder.newInstance());
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet("data");
            Assert.assertEquals("d-1", w.getCellValueAsString(2, 1));
            w.selectSheet("parameters");
            Assert.assertEquals("p.A", w.getCellValueAsString(2, 1));
        }
    }
}
