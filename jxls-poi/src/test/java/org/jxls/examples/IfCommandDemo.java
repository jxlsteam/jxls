package org.jxls.examples;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.common.Context;

public class IfCommandDemo {

    @Test
    public void test() {
        JxlsTester.xls(getClass()).processTemplate(new Context());
    }
}
