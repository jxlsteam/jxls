package org.jxls.templatebasedtests.multisheet;

import static org.junit.Assert.assertEquals;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.junit.Assert;
import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.common.Context;
import org.jxls.common.JxlsException;
import org.jxls.expression.ExpressionEvaluatorFactoryJexlImpl;
import org.jxls.logging.JxlsLogger;
import org.jxls.transform.SafeSheetNameBuilder;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;
import org.jxls.unittests.PoiSafeSheetNameBuilderUnitTest;

/**
 * This is a multi sheet test.
 * 
 * @see PoiSafeSheetNameBuilderUnitTest
 */
public class PoiSafeSheetNameBuilderTest extends AbstractMultiSheetTest {

    /**
     * Tests whether the SafeSheetNameBuilder is used correctly.
     */
    @Test
    public void testSafeSheetNameBuilder() {
        Context context = new Context();
        final List<String> safeNames = new ArrayList<>();
        context.putVar(SafeSheetNameBuilder.CONTEXT_VAR_NAME, new SafeSheetNameBuilder() {
            @Override
            public String createSafeSheetName(String givenSheetName, int index, JxlsLogger logger) {
                String ret = givenSheetName + " sheet";
                safeNames.add(ret);
                return ret;
            }
        });
        List<TestSheet> testSheets = getTestSheets();
        context.putVar("sheets", testSheets);
        context.putVar("sheetnames", getSheetnames(testSheets));
        
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass());
        tester.test(context.toMap(), JxlsPoiTemplateFillerBuilder.newInstance());
        
        assertEquals("Number of safeNames is wrong", 2, safeNames.size());
        assertEquals("Name of 1st sheet is wrong", "data sheet", safeNames.get(0));
        assertEquals("Name of 2nd sheet is wrong", "parameters sheet", safeNames.get(1));
    }
    
    /**
     * Tests whether it still works without a SafeSheetNameBuilder.
     */
    @Test
    public void testNoSafeSheetNameBuilder() throws IOException {
        Context context = new Context();
        List<TestSheet> testSheets = getTestSheets();
        context.putVar("sheets", testSheets);
        context.putVar("sheetnames", getSheetnames(testSheets));
        
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass());
        tester.test(context.toMap(), JxlsPoiTemplateFillerBuilder.newInstance());
    }

    /**
     * Tests what happens if there's no SafeSheetNameBuilder and an invalid sheet name.
     */
    @Test
    public void testNoSafeSheetNameBuilder_invalidName() throws IOException {
        // Prepare
        Context context = new Context();
        List<TestSheet> testSheets = getTestSheets();
        testSheets.get(0).setName("data["); // make name invalid
        context.putVar("sheets", testSheets);
        context.putVar("sheetnames", getSheetnames(testSheets));
        
        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass());
        try {
            tester.test(context.toMap(), JxlsPoiTemplateFillerBuilder.newInstance()
                    .withExceptionThrower()
                    .withExpressionEvaluatorFactory(new ExpressionEvaluatorFactoryJexlImpl(true)));
            
            // Verify
            Assert.fail("JxlsException \"...'data['...\" expected!");
        } catch (JxlsException e) {
            Assert.assertTrue("There must be an ERROR message in the log regarding the invalid sheet name 'data['",
                    e.getMessage().contains("'data['"));
        }
    }
}
