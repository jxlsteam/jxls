package org.jxls.templatebasedtests;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.common.Context;
import org.jxls.util.CannotOpenWorkbookException;

/**
 * Issue B185: Better error message for corrupt Excel file
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
        try (JxlsTester tester = JxlsTester.xlsx(getClass())) {
            tester.processTemplate(new Context());
        }
    }
}
