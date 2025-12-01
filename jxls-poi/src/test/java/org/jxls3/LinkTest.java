package org.jxls3;

import static org.junit.Assert.assertEquals;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.TestWorkbook;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

public class LinkTest {

    @Test
    public void link() throws IOException {
        Map<String, Object> data = new HashMap<>();
        data.put("label1", "go to sheet 2");
        data.put("href", "https://jxls.sourceforge.net/commands.html");
        data.put("label2", "go to website");
        
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass());
        tester.test(data, JxlsPoiTemplateFillerBuilder.newInstance().withExceptionThrower());
        
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet("Sheet1");
            assertEquals("go to sheet 2=Sheet2!B3", w.getHyperlink(3, 2)); // B3
            w.selectSheet("Sheet2");
            assertEquals("go to website=https://jxls.sourceforge.net/commands.html", w.getHyperlink(4, 2)); // B4 
        }
    }
}
