package org.jxls.templatebasedtests;

import java.io.IOException;
import java.util.Arrays;
import java.util.Collection;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.common.Context;

public class Issue127Test {

    @Test
    public void test() throws IOException {
        Collection<Integer> datas = Arrays.asList(1, 2, 3, 4);
        Context context = new Context();
        context.putVar("datas", datas);

        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);
        
        // TODO assertions
    }
}
