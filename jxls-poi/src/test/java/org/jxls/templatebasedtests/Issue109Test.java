package org.jxls.templatebasedtests;

import java.io.IOException;
import java.util.ArrayList;

import org.junit.Ignore;
import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.common.Context;

/**
 * Issue with circular formula
 */
public class Issue109Test {

    @Ignore // TODO MW: not working. Exception! Old class in jxls-examples restored.
    @Test
    public void test() throws IOException {
        Context context = new Context();
        context.putVar("emptyList", new ArrayList<>());

        JxlsTester tester = JxlsTester.xls(getClass());
        tester.processTemplate(context);
        
        // TODO assertions
    }
}
