package org.jxls.templatebasedtests;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.io.IOUtils;
import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.common.Context;

/**
 * Test case for issue#159 Insert image and text underline issues A modification
 * of the original example by ZhengJin Fang
 */
public class Issue159Test {

    @Test
    public void test() throws IOException {
        Context context = new Context();
        Map<String, Object> model = getModel();
        for (String x : model.keySet()) {
            context.putVar(x, model.get(x));
        }

        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);
        
        // TODO assertions
    }

    private Map<String, Object> getModel() throws IOException {
        Map<String, Object> model = new HashMap<>();
        model.put("name", "name111111111");
        model.put("remark", "remark remark remark remark remark remark\n remark remark remark remark remark remark\n remark remark remark remark remark remark remark remark remark remark ");
        
        List<Integer> details = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            details.add(i);
        }
        model.put("details", details);
        
        byte[] stampByteArray;
        try (InputStream stamp = Issue159Test.class.getResourceAsStream("stamp.png")) {
            stampByteArray = IOUtils.toByteArray(stamp);
        }
        model.put("stampImage", stampByteArray);
        return model;
    }
}
