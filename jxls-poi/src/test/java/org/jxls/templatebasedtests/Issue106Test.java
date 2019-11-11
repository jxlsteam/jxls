package org.jxls.templatebasedtests;

import java.io.IOException;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.common.Context;

/**
 * Issue with row height and jxls2 root not in A1
 */
public class Issue106Test {

    @Test
    public void test() throws IOException {
        JxlsTester tester = JxlsTester.xls(getClass());
        tester.processTemplate(new Context());
        
        // TODO assertions
    }
}
