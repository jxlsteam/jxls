package org.jxls.examples;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.common.Context;
import org.jxls.common.ContextImpl;
import org.jxls.entity.Employee;

public class ShiftStrategyDemo {

    @Test
    public void test() {
        Context context = new ContextImpl();
        context.putVar("employees", Employee.generateSampleEmployeeData());
                
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);
    }
    // JXLS team comment: old code was: JxlsHelper.getInstance().processTemplateAtCell(is, os, context, "Result!A1");
}
