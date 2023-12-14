package org.jxls3;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;
import static org.jxls.builder.JxlsStreaming.AUTO_DETECT;
import static org.jxls.builder.JxlsStreaming.STREAMING_OFF;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.TestWorkbook;
import org.jxls.builder.JxlsStreaming;
import org.jxls.entity.Employee;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

public class GroupByTest {

    @Test
    public void asc_standard() throws IOException {
        asc(STREAMING_OFF);
    }

    @Test
    public void asc_streaming() throws IOException {
        asc(AUTO_DETECT);
    }
    
    private void asc(JxlsStreaming streaming) {
    	// Template also used by ExpressionEvaluatorFactoryTest.groupBy()
        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass(), "asc");
        tester.test(data(), JxlsPoiTemplateFillerBuilder.newInstance().withStreaming(streaming));
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet(0);
            assertEquals("high", w.getCellValueAsString(2, 1)); // A2
            assertTrue(w.getCellValueAsDouble(3, 3) > 2000); // C3
            assertTrue(w.getCellValueAsDouble(4, 3) > 2000);
            assertTrue(w.getCellValueAsDouble(5, 3) > 2000);
            assertEquals("normal", w.getCellValueAsString(6, 1)); // A6
            assertTrue(w.getCellValueAsDouble(7, 3) <= 2000);
            assertTrue(w.getCellValueAsDouble(8, 3) <= 2000);
        }
    }

    @Test
    public void desc() {
        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass(), "desc");
        tester.test(data(), JxlsPoiTemplateFillerBuilder.newInstance());
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet(0);
            assertEquals("normal", w.getCellValueAsString(2, 1)); // A2
            assertTrue(w.getCellValueAsDouble(3, 3) <= 2000); // C3
            assertTrue(w.getCellValueAsDouble(4, 3) <= 2000);
            assertEquals("high", w.getCellValueAsString(5, 1)); // A5
            assertTrue(w.getCellValueAsDouble(6, 3) > 2000);
            assertTrue(w.getCellValueAsDouble(7, 3) > 2000);
            assertTrue(w.getCellValueAsDouble(8, 3) > 2000);
        }
    }

    @Test
    public void select() {
        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass(), "select");
        tester.test(data(), JxlsPoiTemplateFillerBuilder.newInstance());
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet(0);
            assertEquals("high", w.getCellValueAsString(2, 1)); // A2
            assertTrue(w.getCellValueAsDouble(3, 3) > 2000 && w.getCellValueAsDouble(3, 3) < 2500); // C3
            assertEquals("normal", w.getCellValueAsString(4, 1)); // A4
            assertTrue(w.getCellValueAsDouble(5, 3) <= 2000); // C5
        }
    }

    private Map<String, Object> data() {
        Map<String, Object> data = new HashMap<>();
        data.put("employees", Employee.generateSampleEmployeeData());
        return data;
    }
}
