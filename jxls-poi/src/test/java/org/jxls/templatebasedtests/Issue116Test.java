package org.jxls.templatebasedtests;

import java.io.IOException;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.common.Context;

/**
 * Formula handling issues (formula external to any jx:area)
 */
public class Issue116Test {

    @Test
    public void externalFormulas() throws IOException {
        Context context = new Context();
        context.putVar("vars", Arrays.asList(1.234, 5.678, 3.1234, 8.9090, 12.34567));
        
        JxlsTester tester = JxlsTester.xlsx(getClass(), "externalFormulas");
        tester.setUseFastFormulaProcessor(true);
        tester.processTemplate(context);
        
        // TODO assertions
    }

    @Test
    public void forEachFormulas() throws IOException {
        Context context = new Context();
        context.putVar("vars", Arrays.asList(1.234, 5.678, 3.1234, 8.9090, 12.34567));
        Map<String, Object> jxlsMap = new HashMap<>();
        jxlsMap.put("math", Math.class); // add Math utility functions

        JxlsTester tester = JxlsTester.xlsx(getClass(), "forEachFormulas");
        tester.processTemplate(context);
    }
}