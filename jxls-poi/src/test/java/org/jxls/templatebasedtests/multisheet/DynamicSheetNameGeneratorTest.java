package org.jxls.templatebasedtests.multisheet;

import org.junit.Test;
import org.jxls.JxlsTester;
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
    public void testWithExpression() {
        Context context = new Context();
        context.putVar("sheets", getTestSheets());
        
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);
        
        // Result: Sheets with name "data" and "parameters" created.
    }
}
