package org.jxls.examples;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.common.Context;
import org.jxls.common.ContextImpl;
import org.jxls.entity.Employee;

public class ParameterizedFormulasDemo {

    @Test
    public void test() {
        Context context = new ContextImpl();
        context.putVar("employees", Employee.generateSampleEmployeeData());
        context.putVar("bonus", 0.1);

        JxlsTester tester = JxlsTester.xls(getClass());
        tester.processTemplate(context);
    }
    // JXLS team comment: old code was: JxlsHelper.getInstance().processTemplateAtCell(is, os, context, "Result!A1");
}
