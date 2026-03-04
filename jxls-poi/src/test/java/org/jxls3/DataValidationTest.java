package org.jxls3;

import java.util.HashMap;
import java.util.Map;

import org.junit.Assert;
import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.TestWorkbook;
import org.jxls.builder.JxlsStreaming;
import org.jxls.templatebasedtests.IssueB133Test;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

public class DataValidationTest {

    @Test
    public void streamingOff() {
        check(JxlsStreaming.STREAMING_OFF);
    }

    @Test
    public void streamingOn() {
        check(JxlsStreaming.STREAMING_ON);
    }

    private void check(JxlsStreaming streaming) {
        // Prepare
        Map<String, Object> data = new HashMap<>();
        data.put("employees", IssueB133Test.createEmployees());
        
        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass());
        tester.test(data, JxlsPoiTemplateFillerBuilder.newInstance().withStreaming(streaming));
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            String x = "|LIST|F1=\"AA,BB,CCCC\"|LZE=true|DA=true\n";
            w.selectSheet("down");
            Assert.assertEquals("down test failed", "C3_C4:C6_".replace("_", x), w.getDataValidations());
            
            x = "|LIST|F1=\"F,G,H,JJJ\"|LZE=true|DA=true\n";
            w.selectSheet("right");
            Assert.assertEquals("right test failed", "B1_D1_F1_H1_".replace("_", x), w.getDataValidations());
        }
    }
}
