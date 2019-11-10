package org.jxls.templatebasedtests;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

import org.junit.Test;
import org.jxls.common.Context;
import org.jxls.util.CannotOpenWorkbookException;
import org.jxls.util.JxlsHelper;

/**
 * Issue 185 - Better error message for corrupt Excel file
 * 
 * <pre> Old error message: java.lang.IllegalStateException: Cannot load XLS transformer. Please make sure a Transformer implementation is in classpath
 * at org.jxls.util.JxlsHelper.createTransformer(JxlsHelper.java:407)
 * at org.jxls.util.JxlsHelper.processTemplate(JxlsHelper.java:186)
 * 
 * New error message: org.jxls.util.CannotOpenWorkbookException: java.io.IOException: Your InputStream was neither an OLE2 stream, nor an OOXML stream</pre>
 */
public class CorruptExcelFileTest {

    @Test(expected = CannotOpenWorkbookException.class)
    public void test() throws Exception {
        String out = "target/CorruptExcelFileTest_output.xlsx";
        try (InputStream is = getClass().getResourceAsStream("CorruptExcelFileTest.xlsx")) { // corrupt Excel file (It's a .png file.)
            try (OutputStream os = new FileOutputStream(out)) {
                JxlsHelper.getInstance().processTemplate(is, os, new Context());
            }
        } finally {
            new File(out).delete();
        }
    }
}
