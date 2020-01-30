package org.jxls.examples;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.common.Context;
import org.jxls.entity.Department;

/**
 * Created by Leonid Vysochyn on 6/30/2015.
 * todo: improve each command to be able to set merge cells
 */
public class MergedCellsDemo  {
    
    @Test
    public void test() {
        Context context = new Context();
        context.putVar("departments", Department.createDepartments());
        
        JxlsTester tester = JxlsTester.xls(getClass());
        tester.processTemplate(context);
    }
}
