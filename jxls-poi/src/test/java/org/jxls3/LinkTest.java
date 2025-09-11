package org.jxls3;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

public class LinkTest {

    @Test
    public void link() throws IOException {
        Map<String, Object> data = new HashMap<>();
        data.put("href", "https://jxls.sourceforge.net/commands.html");
        data.put("label", "Commands");
        Jxls3Tester.xlsx(getClass()).test(data, JxlsPoiTemplateFillerBuilder.newInstance().withExceptionThrower());
        // open file to verify
    }
}
