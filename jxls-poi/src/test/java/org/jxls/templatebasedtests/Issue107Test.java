package org.jxls.templatebasedtests;

import java.io.IOException;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.common.Context;

/**
 * Issue with formatting of parts of text
 */
public class Issue107Test {

    @Test
    public void test() throws IOException {
        JxlsTester tester = JxlsTester.xls(getClass());
        tester.processTemplate(new Context());
        
        // TODO assertions
    }
}
