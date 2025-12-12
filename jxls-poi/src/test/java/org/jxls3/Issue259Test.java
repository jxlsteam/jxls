package org.jxls3;

import static org.junit.Assert.assertTrue;

import java.util.HashMap;
import java.util.List;

import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.TestWorkbook;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

/**
 * Merged cell region outside jx:area is missing
 */
public class Issue259Test {

    @Test
    public void test() {
        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass());
        tester.test(new HashMap<>(), JxlsPoiTemplateFillerBuilder.newInstance());

        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet(0);
            List<CellRangeAddress> mergedRegions = w.getMergedRegions();
            assertTrue("Missing merged region B5:C5", mergedRegions.stream().anyMatch(i -> "B5:C5".equals(i.formatAsString())));
            assertTrue("Missing merged region A1:C3", mergedRegions.stream().anyMatch(i -> "A1:C3".equals(i.formatAsString())));
        }
    }
}
