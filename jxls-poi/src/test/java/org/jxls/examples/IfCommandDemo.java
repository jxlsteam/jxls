package org.jxls.examples;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.common.ContextImpl;

public class IfCommandDemo {

    @Test
    public void test() {
        JxlsTester.xlsx(getClass()).processTemplate(new ContextImpl());
    }
}
