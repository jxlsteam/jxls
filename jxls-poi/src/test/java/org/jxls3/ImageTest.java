package org.jxls3;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.command.ImageCommand;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

public class ImageTest {

    @Test
    public void png() throws IOException {
        Map<String, Object> data = new HashMap<>();
        data.put("image", ImageCommand.toByteArray(ImageTest.class.getResourceAsStream("/org/jxls/examples/business.png")));
        Jxls3Tester.xlsx(getClass()).test(data, JxlsPoiTemplateFillerBuilder.newInstance().withExceptionThrower());
        // open file to verify
    }
}
