package org.jxls.examples;

import java.util.HashMap;
import java.util.Map;

import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.entity.Org;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

/**
 * @author Leonid Vysochyn
 */
public class FormulaCopyDemo {

    @Test
    public void test() {
        Map<String,Object> data = new HashMap<>();
        data.put("orgs", Org.generate(3, 3));
        Jxls3Tester.xlsx(getClass()).test(data, JxlsPoiTemplateFillerBuilder.newInstance());
    }
}
