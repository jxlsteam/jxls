package org.jxls.templatebasedtests;

import java.io.IOException;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.common.Context;

/**
 * XLS file that has sheet with chart is not being processed
 */
public class Issue119Test {

    @Test
    public void test() throws IOException {
        Context context = new Context();
        context.putVar("title", "Report XLS");
        context.putVar("value", 100);

        JxlsTester tester = JxlsTester.xls(getClass());
        tester.processTemplate(context);
        
        // TODO assertions -> Ticket state is WONT-FIX. Delete testcase?
    }
}
