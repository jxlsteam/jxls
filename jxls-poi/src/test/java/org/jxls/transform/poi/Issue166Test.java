package org.jxls.transform.poi;

import org.junit.Test;
import org.jxls.command.TestWorkbook;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.junit.Assert.assertEquals;

/**
 * Issue 166: Wrong average on 2nd sheet
 */
public class Issue166Test {
    private final static String INPUT_FILE_PATH = "issue166_template.xlsx";
    private final static String OUTPUT_FILE_PATH = "target/issue166_output.xlsx";

    @Test
    public void testFormulasOnBothSheets() throws IOException {
        // define result set
        List<Map<String, Object>> rs = new ArrayList<>();
        for (int i = 0; i < 5; i++) {
            Map<String, Object> map = new HashMap<String, Object>();
            map.put("count", i);
            rs.add(map);
        }
        // process templates
        File out = new File(OUTPUT_FILE_PATH);
        try(InputStream is = Issue166Test.class.getResourceAsStream(INPUT_FILE_PATH)) {
            try (OutputStream os = new FileOutputStream(out)) {
                JxlsHelper jxlsHelper = JxlsHelper.getInstance();
                Context context = new Context();
                context.putVar("rs0", rs);
                jxlsHelper.setEvaluateFormulas(false);
                jxlsHelper.processTemplate(is, os, context);
            }
        }

        // verify result
        try (TestWorkbook w = new TestWorkbook(out)) {
            verifyTab(w, "Tab1");
            verifyTab(w, "Tab2");
        }
    }

    private void verifyTab(TestWorkbook w, String tabName) {
        w.selectSheet(tabName);
        assertEquals(1, w.getCellValueAsDouble(3, 1), 1e-3);
        assertEquals(4, w.getCellValueAsDouble(6, 1), 1e-3);
        assertEquals("AVERAGEA(A2:A6)", w.getFormulaString(7, 2));
        assertEquals("SUM(A2:A6)", w.getFormulaString(8, 2));
    }

}
