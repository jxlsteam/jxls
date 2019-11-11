package org.jxls.templatebasedtests;

import java.util.ArrayList;
import java.util.Arrays;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.common.Context;

public class Issue103Test {

    @Test
    public void test() {
        Context context = new Context();
        context.putVar("nonEmptyList", new ArrayList<>(Arrays.asList("A", "B", "C", "D", "E", "F", "G", "H", "I", "J")));
        context.putVar("emptyList", new ArrayList<String>());
        
        JxlsTester tester = JxlsTester.xls(getClass());
        tester.processTemplate(context);
        
        // TODO assertions
    }
}
