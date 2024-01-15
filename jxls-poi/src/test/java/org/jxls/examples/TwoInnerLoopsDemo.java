package org.jxls.examples;

import java.io.IOException;
import java.text.ParseException;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.common.Context;
import org.jxls.common.ContextImpl;
import org.jxls.entity.Department;

/**
 * Two nested each commands demo
 */
public class TwoInnerLoopsDemo {

    @Test
    public void test() throws ParseException, IOException {
        Context context = new ContextImpl();
        context.putVar("departments", Department.createDepartments());
        
        JxlsTester tester = JxlsTester.xls(getClass());
        tester.processTemplate(context);
    }
}
