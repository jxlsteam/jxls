package org.jxls.templatebasedtests;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.BeforeClass;
import org.junit.Test;
import org.jxls.common.Context;
import org.jxls.transform.Transformer;
import org.jxls.transform.poi.SelectSheetsForStreamingPoiTransformer;
import org.jxls.util.JxlsHelper;

public class TableTest {
    private static final File dir = new File("test-output");
    
    @BeforeClass
    public static void init() {
        dir.mkdirs();
    }
    
    @Test
    public void testTable() throws IOException {
        checkTable(101);
        // Please test here, if the table has been extended.
    }

    @Test
    public void testEmptyTable() throws IOException {
        // This 'empty table' testcase uses the standard PoiTransformer.
        checkTable(0);
        // Here you can test if you can open the result file.
    }

    private void checkTable(int max) throws IOException {
        Context ctx = createContextWithTestData(max);
        InputStream in = TableTest.class.getResourceAsStream("table.xlsx");
        try {
            FileOutputStream out = new FileOutputStream(new File(dir, "table_" + max + ".xlsx"));
            try {
                Transformer transformer = JxlsHelper.getInstance().createTransformer(in, out);
                JxlsHelper.getInstance().processTemplate(ctx, transformer);
            } finally {
                out.close();
            }
        } finally {
            in.close();
        }
    }

    /**
     * With streaming this testcase takes less than 5 seconds. Without streaming it takes much more time.
     */
    @Test
    public void testTableWithStreaming() throws IOException {
        checkTableWithStreaming(100000);
        // Please test if the file has been created fast and low in memory.
    }

    @Test
    public void testEmptyTableWithStreaming() throws IOException {
        // This 'empty table' testcase uses a special streaming transformer.
        checkTableWithStreaming(0);
        // Here you can test if you can open the result file.
    }

    /**
     * Use streaming only for the sheet named "table".
     */
    private void checkTableWithStreaming(int max) throws IOException {
        Context ctx = createContextWithTestData(max);
        InputStream in = TableTest.class.getResourceAsStream("table.xlsx");
        try {
            FileOutputStream out = new FileOutputStream(new File(dir, "table_" + max + "_streaming.xlsx"));
            try {
                Workbook workbook = WorkbookFactory.create(in);
                SelectSheetsForStreamingPoiTransformer transformer = new SelectSheetsForStreamingPoiTransformer(workbook);
                Set<String> streamedSheets = new HashSet<String>();
                streamedSheets.add("table");
                transformer.setDataSheetsToUseStreaming(streamedSheets);
                transformer.setOutputStream(out);
                JxlsHelper.getInstance().processTemplate(ctx, transformer);
            } finally {
                out.close();
            }
        } finally {
            in.close();
        }
    }

    private Context createContextWithTestData(int max) {
        List<TableTestObject> list = new ArrayList<TableTestObject>();
        List<TableTestObject> list20 = new ArrayList<TableTestObject>();
        for (int i = 1; i <= max; i++) {
            TableTestObject a = new TableTestObject("name-" + i, "address-" + i);
            list.add(a);
            if (i <= 20) {
                list20.add(a);
            }
        }
        Context ctx = new Context();
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
