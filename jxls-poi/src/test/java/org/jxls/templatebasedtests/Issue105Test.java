package org.jxls.templatebasedtests;

import java.io.IOException;
import java.util.Arrays;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.common.Context;

/**
 * Issue with 'big double values' (like 1.3E22) being parsed as cell references
 * (like E22)
 */
public class Issue105Test {

    @Test
    public void test() throws IOException {
        Context context = new Context();
        context.putVar("vars", Arrays.asList(new Values(), new Values(), new Values()));

        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);

        // TODO assertions
    }

    public static class Values {
        public double smallValue = 1.2;
        public double bigValue = 1.3E22;
    }
}