package org.jxls.templatebasedtests;

import java.util.ArrayList;
import java.util.List;

import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.JxlsTester;
import org.jxls.builder.JxlsStreaming;
import org.jxls.common.Context;
import org.jxls.common.ContextImpl;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

// TODO MW: missing assertions (can file be opened? record size okay? speed okay?)
public class TableTest {
    
    @Test
    public void testTable() {
        checkTable(101);
        // Please test here, if the table has been extended.
    }

    @Test
    public void testEmptyTable() {
        // This 'empty table' testcase uses the standard PoiTransformer.
        checkTable(0);
        // Here you can test if you can open the result file.
    }

    private void checkTable(int max) {
        Context context = createContextWithTestData(max);
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);
    }

    /**
     * With streaming this testcase takes less than 5 seconds. Without streaming it takes much more time.
     */
    @Test
    public void testTableWithStreaming() {
        checkTableWithStreaming(100000);
        // Please test if the file has been created fast and low in memory.
    }

    @Test
    public void testEmptyTableWithStreaming() {
        // This 'empty table' testcase uses a special streaming transformer.
        checkTableWithStreaming(0);
        // Here you can test if you can open the result file.
    }

    /**
     * Use streaming only for the sheet named "table".
     */
    private void checkTableWithStreaming(int max) {
        Context context = createContextWithTestData(max);
        
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass());
        tester.test(context.toMap(), JxlsPoiTemplateFillerBuilder.newInstance().withStreaming(JxlsStreaming.AUTO_DETECT));
    }

    private Context createContextWithTestData(int max) {
        List<TableTestObject> list = new ArrayList<>();
        List<TableTestObject> list20 = new ArrayList<>();
        for (int i = 1; i <= max; i++) {
            TableTestObject a = new TableTestObject("name-" + i, "address-" + i);
            list.add(a);
            if (i <= 20) {
                list20.add(a);
            }
        }
        Context ctx = new ContextImpl();
        ctx.putVar("list", list);
        ctx.putVar("list20", list20); // a short list for the 2nd sheet; always without streaming
        return ctx;
    }
    
    public static class TableTestObject {
        private final String name;
        private final String address;

        public TableTestObject(String name, String address) {
            this.name = name;
            this.address = address;
        }

        public String getName() {
            return name;
        }

        public String getAddress() {
            return address;
        }
    }
}
