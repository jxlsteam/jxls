package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;

import java.io.IOException;
import java.util.Arrays;
import java.util.Collection;

import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;
import org.jxls.common.ContextImpl;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

/**
 * Problem with formulas referencing cells in another worksheet containing special characters in name
 */
public class IssueB127Test {

    @Test
    public void test() throws IOException {
        // Prepare
        Collection<Integer> datas = Arrays.asList(1, 2, 3, 4);
        Context context = new ContextImpl();
        context.putVar("datas", datas);

        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass());
        tester.test(context.toMap(), JxlsPoiTemplateFillerBuilder.newInstance());
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet("OK");
            assertEquals(1d, w.getCellValueAsDouble(3, 1), 0.1d);
            assertEquals(2d, w.getCellValueAsDouble(4, 1), 0.1d);
            assertEquals(3d, w.getCellValueAsDouble(5, 1), 0.1d);
            assertEquals(4d, w.getCellValueAsDouble(6, 1), 0.1d);

            w.selectSheet("KO_");
            assertEquals(1d, w.getCellValueAsDouble(3, 1), 0.1d);
            assertEquals(2d, w.getCellValueAsDouble(4, 1), 0.1d);
            assertEquals(3d, w.getCellValueAsDouble(5, 1), 0.1d);
            assertEquals(4d, w.getCellValueAsDouble(6, 1), 0.1d);

            w.selectSheet("Formulas");
            assertEquals(10d, w.getCellValueAsDouble(2, 1), 0.1d);
            assertEquals(10d, w.getCellValueAsDouble(2, 2), 0.1d);
        }
    }
}
