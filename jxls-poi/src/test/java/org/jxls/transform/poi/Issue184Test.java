package org.jxls.transform.poi;

import static org.junit.Assert.assertEquals;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.junit.Test;
import org.jxls.command.TestWorkbook;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

/**
 * Issue 184
 * 
 * The template has a jx:if inside a jx:each. The sums of the columns before the jx:if-column are wrong.
 * 
 * Cause:
 * commit 5354beaf6c73bd252f98d272ac1e4dc8f870a13b,
 * issue #160 - add isForwardOnly() method to Transformer and processing static cells differently depending on this flag
 */
public class Issue184Test {

    @Test
    public void test() throws Exception {
        // Prepare
        String in = "issue184_template.xls";
        File out = new File("target/issue184_output.xls");
        List<Integer> data = new ArrayList<>();
        for (int j = 1; j <= 3; j++) {
            data.add(j);
        }
        Context context = new Context();
        context.putVar("data", data);
        context.putVar("b", 100);
        
        // Test
        try (InputStream is = getClass().getResource(in).openStream()) {
            try (FileOutputStream os = new FileOutputStream(out)) {
                JxlsHelper.getInstance().setEvaluateFormulas(true).processTemplate(is, os, context);
            }
        }
        
        // Verify
        try (TestWorkbook w = new TestWorkbook(out)) {
            w.selectSheet("Bug");
            assertEquals(6d, w.getCellValueAsDouble(5, 3), 0.01d);
            assertEquals(6d, w.getCellValueAsDouble(5, 4), 0.01d);
            assertEquals(6d, w.getCellValueAsDouble(5, 5), 0.01d);
            assertEquals(6d, w.getCellValueAsDouble(5, 1), 0.01d);
            assertEquals(6d, w.getCellValueAsDouble(5, 2), 0.01d);
            assertEquals("SUM(C2:C4)", w.getFormulaString(5, 3));
        }
    }
}
