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

import static org.junit.Assert.assertEquals;

/**
 * A test case for Issue 188 Referencing other sheet in JXLS-processed cell formula replaces formula with "=0"
 */
public class Issue188Test {
    private final static String INPUT_FILE_PATH = "issue188_template.xlsx";
    private final static String OUTPUT_FILE_PATH = "target/issue188_output.xlsx";

    @Test
    public void testCrossSheetFormulas() throws IOException {
        // prepare
        File out = new File(OUTPUT_FILE_PATH);
        try(InputStream is = Issue188Test.class.getResourceAsStream(INPUT_FILE_PATH)) {
            try (OutputStream os = new FileOutputStream(out)) {
                Context context = new Context();
                JxlsHelper.getInstance().setEvaluateFormulas(true).processTemplate(is, os, context);
            }
        }
        // verify
        try (TestWorkbook w = new TestWorkbook(out)) {
            w.selectSheet("Sheet1");
            assertEquals("Sheet2!A1", w.getFormulaString(1, 1));
            assertEquals("Static data", w.getCellValueAsString(1, 1));
        }

    }
}
