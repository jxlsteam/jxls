package org.jxls.templatebasedtests;

import java.io.IOException;
import java.util.Date;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.common.Context;
import org.jxls.entity.Employee;

/**
 * Excel templates context variable not found if used beyond column AZ
 */
public class Issue167Test {

    @Test
    public void test() throws IOException {
        Context context = new Context();
        context.putVar("var1", "Variable 1");
        context.putVar("var2", new Employee("Leo", new Date(), 1000, 0.10));

        JxlsTester tester = JxlsTester.xls(getClass());
        tester.processTemplate(context);
        
        // TODO assertions
    }
}
