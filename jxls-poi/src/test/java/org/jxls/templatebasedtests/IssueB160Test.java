package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;
import static org.jxls.builder.JxlsStreaming.STREAMING_ON;

import java.io.IOException;
import java.text.ParseException;
import java.util.Arrays;
import java.util.Collection;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;
import org.jxls.transform.poi.PoiContext;

/**
 * jx:each with direction=RIGHT with SXSSF Transformer rewrites static cells
 */
public class IssueB160Test {

    @Test
    public void test() throws IOException, ParseException {
        // Prepare
        Context context = new PoiContext();
        context.putVar("lotsOfStuff", createLotsOfStuff());
        context.putVar("columns", new Columns());
        
        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass());
        tester.test(context.toMap(), JxlsPoiTemplateFillerBuilder.newInstance().withStreaming(STREAMING_ON.withOptions(2, false, false)));

        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet(0);
            assertEquals("header2_dynamic", w.getCellValueAsString(4, 3));
            assertEquals("stuff_2_value2", w.getCellValueAsString(6, 3));
            assertEquals("Static last line", w.getCellValueAsString(7, 1));
        }
    }

    private static List<Map<String, Object>> createLotsOfStuff() {
        Map<String, Object> stuff1 = new LinkedHashMap<>();
        Map<String, Object> stuff2 = new LinkedHashMap<>();

        stuff1.put("header0", "stuff_1_value0");
        stuff1.put("header1_dynamic", "stuff_1_value1");
        stuff1.put("header2_dynamic", "stuff_1_value2");
        stuff1.put("header3_dynamic", "stuff_1_value3");

        stuff2.put("header0", "stuff_2_value0");
        stuff2.put("header1_dynamic", "stuff_2_value1");
        stuff2.put("header2_dynamic", "stuff_2_value2");
        stuff2.put("header3_dynamic", "stuff_2_value3");

        return Arrays.asList(stuff1, stuff2);
    }
    
    public static class Columns {
        
        public Collection<String> keyOf(List<Map<String, Object>> row) {
            return row.get(0).keySet().stream().filter(k -> k.endsWith("_dynamic")).collect(Collectors.toList());
        }

        public Collection<Object> valueOf(Map<String, Object> row) {
            return row.entrySet().stream()
                    .filter(entry -> entry.getKey() != null && entry.getKey().endsWith("_dynamic"))
                    .map(Map.Entry::getValue)
                    .collect(Collectors.toList());
        }
    }
}
