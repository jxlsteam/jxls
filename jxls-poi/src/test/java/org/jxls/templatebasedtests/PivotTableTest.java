package org.jxls.transform.poi;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.junit.Test;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

public class PivotTableTest {

    /**
     * issue 155: Pivot table does not work with NLS
     */
    @Test
    public void nls() throws IOException {
        InputStream in = PivotTableTest.class.getResourceAsStream("pivottable.xlsx");
        File outputFile = new File("target/pivottable.xlsx");
        FileOutputStream out = new FileOutputStream(outputFile);
        Context context = new Context();
        context.putVar("R", getResources()); // NLS
        context.putVar("list", getTestData());
        TestPTTransformer transformer = new TestPTTransformer(PoiTransformer.createTransformer(in, out), context);
        JxlsHelper.getInstance().processTemplate(context, transformer);
        
        System.out.println(outputFile.getAbsolutePath()); // XXX
        // result: broken PivotTable
    }

    private Map<String, String> getResources() {
        Map<String, String> r = new HashMap<>();
        r.put("name", "Name (EN)");
        r.put("city", "City (EN)");
        return r;
    }

    private List<Map<String, String>> getTestData() {
        List<Map<String, String>> list = new ArrayList<>();
        add(list, "Leonid", "Danzig");
        add(list, "Heil", "Berlin");
        add(list, "Marcus", "Krefeld");
        add(list, "Merkel", "Berlin");
        add(list, "Seehofer", "Berlin");
        add(list, "Waldemar", "Krefeld");
        return list;
    }
    
    private void add(List<Map<String, String>> list, String name, String city) {
        Map<String, String> map = new HashMap<>();
        map.put("name", name);
        map.put("city", city);
        list.add(map);
    }
}
