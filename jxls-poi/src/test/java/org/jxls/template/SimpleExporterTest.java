package org.jxls.template;

import static org.junit.Assert.assertEquals;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.junit.Test;
import org.jxls.TestWorkbook;

public class SimpleExporterTest {

    @Test
    public void test() throws FileNotFoundException {
        List<List<Object>> list = new ArrayList<>();
        List<Object> item = new ArrayList<>();
        item.add("abc");
        item.add(Double.valueOf(17.4));
        list.add(item);
        item = new ArrayList<>();
        item.add("bcd");
        item.add(Double.valueOf(0));
        list.add(item);
        File out = new File("target/SimpleExporterTest_output.xlsx");
        out.getParentFile().mkdirs();
        new SimpleExporter().gridExport(Arrays.asList("A", "B"), list, "Double:B2",
                new FileOutputStream(out));
        
        try (TestWorkbook w = new TestWorkbook(out)) {
            w.selectSheet("Sheet1");
            assertEquals("A", w.getCellValueAsString(1, 1));
            assertEquals("B", w.getCellValueAsString(1, 2));
            assertEquals("abc", w.getCellValueAsString(2, 1));
            assertEquals(Double.valueOf(17.4), w.getCellValueAsDouble(2, 2), 0.005d);
            assertEquals("bcd", w.getCellValueAsString(3, 1));
            assertEquals(Double.valueOf(0), w.getCellValueAsDouble(3, 2), 0.005d);
        }
    }
}
