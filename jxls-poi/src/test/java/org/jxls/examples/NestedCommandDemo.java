package org.jxls.examples;

import org.junit.Test;
import org.jxls.JxlsTester;

/**
 * jx:each command contains nested jx:if command
 */
public class NestedCommandDemo {

    @Test
    public void test() {
        JxlsTester.quickProcessXlsTemplate(getClass());
    }
    // JXLS team comment: old code just processed Result!A1    JxlsHelper.getInstance().processTemplateAtCell(is, os, context, "Result!A1");
}