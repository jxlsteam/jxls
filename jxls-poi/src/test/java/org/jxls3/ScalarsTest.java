package org.jxls3;

import static org.junit.Assert.assertEquals;

import java.io.IOException;
import java.util.HashMap;

import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.TestWorkbook;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

// #360
public class ScalarsTest {

    @Test
    public void testNumberArrays() throws IOException {
    	// Prepare
        var data = new HashMap<String, Object>();
        
        int[] ints = { 0, 1, Integer.MAX_VALUE };
        data.put("ints", ints);

        double[] doubles = { 0, -12345.6d };
        data.put("doubles", doubles);

        long[] longs = { -5000, 2147483648l };
        data.put("longs", longs);
        
        // Test
        var tester = Jxls3Tester.xlsx(getClass());
        tester.test(data, JxlsPoiTemplateFillerBuilder.newInstance());

        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
        	w.selectSheet(0);
        	assertEquals(ints[0], w.getCellValueAsDouble(3, 1), .5); // can read only as double not as int
        	assertEquals(ints[1], w.getCellValueAsDouble(4, 1), .5);
        	assertEquals(ints[2], w.getCellValueAsDouble(5, 1), .5);
        	assertEquals(doubles[0], w.getCellValueAsDouble(8, 1), .5);
        	assertEquals(doubles[1], w.getCellValueAsDouble(9, 1), .5);
        	assertEquals(longs[0], w.getCellValueAsDouble(12, 1), .5);
        	assertEquals(longs[1], w.getCellValueAsDouble(13, 1), .5);
        }
    }
}
