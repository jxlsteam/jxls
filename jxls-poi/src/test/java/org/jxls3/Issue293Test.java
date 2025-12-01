package org.jxls3;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

public class Issue293Test {
    private static int COUNT = 600; // or 1000
    
    /**
     * If the template file has mergedRegions, it is really slow. And will be slower
     * and slower.
     */
    @Test
    @org.junit.Ignore
    public void slow() {
        List<Map<String, String>> mapList = new ArrayList<>();
        for (int i = 0; i < COUNT; i++) {
            Map<String, String> map = new HashMap<>();
            map.put("mergeA", "mergeA" + i);
            map.put("mergeB", "mergeB" + i);
            map.put("mergeC", "mergeC" + i);
            map.put("d", "d" + i);
            map.put("e", "e" + i);
            mapList.add(map);
        }
        HashMap<String, Object> map = new HashMap<>();
        map.put("rowList", mapList);
        
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass(), "slow");
        tester.test(map, JxlsPoiTemplateFillerBuilder.newInstance());
    }

    /**
     * It will be normal when there is no mergedRegions.
     */
    @Test
    public void fast() {
        List<Map<String, String>> mapList = new ArrayList<>();
        for (int i = 0; i < COUNT; i++) {
            Map<String, String> map = new HashMap<>();
            map.put("a", "a" + i);
            map.put("b", "b" + i);
            map.put("c", "c" + i);
            map.put("d", "d" + i);
            map.put("e", "e" + i);
            mapList.add(map);
        }
        HashMap<String, Object> map = new HashMap<>();
        map.put("rowList", mapList);

        Jxls3Tester tester = Jxls3Tester.xlsx(getClass(), "fast");
        tester.test(map, JxlsPoiTemplateFillerBuilder.newInstance());
    }
}
