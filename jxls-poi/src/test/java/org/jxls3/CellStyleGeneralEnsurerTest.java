package org.jxls3;

import static org.junit.Assert.assertEquals;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.TestWorkbook;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

/**
 * Test POI bug 66679 compensation
 */
public class CellStyleGeneralEnsurerTest {
    // CellStyleGeneralEnsurerTest.xlsx has no CellStyle "General" at cell D2.

    @Test
    public void test() throws IOException {
        check(JxlsPoiTemplateFillerBuilder.newInstance().withCellStyleGeneralEnsurer(), "General");
        // Number 2 would be formatted as "2", as expected.
    }

    @Test
    public void negativetest() throws IOException {
        check(JxlsPoiTemplateFillerBuilder.newInstance(), "0.00");
        // Number 2 would be formatted as "2.00", which is not expected.
    }

    private void check(JxlsPoiTemplateFillerBuilder builder, String expected) {
        // Prepare
        Map<String, Object> data = new HashMap<>();
        data.put("n", 2d);
        
        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass());
        tester.test(data, builder);

        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet(0);
            assertEquals(expected, w.getCellDataFormat(2, 4)); // D2
        }
    }
}
