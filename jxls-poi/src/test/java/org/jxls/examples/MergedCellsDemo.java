package org.jxls.examples;

import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.common.Context;
import org.jxls.common.ContextImpl;
import org.jxls.entity.Department;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

/**
 * Created by Leonid Vysochyn on 6/30/2015.
 * todo: improve each command to be able to set merge cells
 */
public class MergedCellsDemo  {
    
    @Test
    public void test() {
        Context context = new ContextImpl();
        context.putVar("departments", Department.createDepartments());
        
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass());
        tester.test(context.toMap(), JxlsPoiTemplateFillerBuilder.newInstance());
    }
}
