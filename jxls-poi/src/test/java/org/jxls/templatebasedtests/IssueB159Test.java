package org.jxls.templatebasedtests;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.io.IOUtils;
import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.common.Context;
import org.jxls.common.ContextImpl;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

/**
 * Testcase for issues 159 (insert image) and 178 (text underline issues)
 * A modification of the original example by ZhengJin Fang
 */
public class IssueB159Test {

    @Test
    public void test() throws IOException {
        // Prepare
        Context context = new ContextImpl();
        Map<String, Object> model = getModel();
        for (String key : model.keySet()) {
            context.putVar(key, model.get(key));
        }

        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass());
        tester.test(model, JxlsPoiTemplateFillerBuilder.newInstance());
        
        // Verify
        // TODO I guess in the template file the image command in cell I9 must be extended with: lockRange=false
        // TODO assertions: What have to be verified?
    }

    private Map<String, Object> getModel() throws IOException {
        // issue 178 >>
        Map<String, Object> model = new HashMap<>();
        model.put("name", "name111111111");
        model.put("remark", "remark remark remark remark remark remark\n remark remark remark remark remark remark\n remark remark remark remark remark remark remark remark remark remark ");
        // <<
        
        List<Integer> details = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            details.add(Integer.valueOf(i));
        }
        model.put("details", details);
        
        // issue 159 numbers 1 and 2 >>
        byte[] stampByteArray;
        try (InputStream stamp = IssueB159Test.class.getResourceAsStream("stamp.png")) {
            stampByteArray = IOUtils.toByteArray(stamp);
        }
        model.put("stampImage", stampByteArray);
        // <<
        return model;
    }
}
