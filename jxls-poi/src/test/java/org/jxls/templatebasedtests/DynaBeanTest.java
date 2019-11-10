package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.beanutils.BasicDynaClass;
import org.apache.commons.beanutils.DynaBean;
import org.apache.commons.beanutils.DynaClass;
import org.apache.commons.beanutils.DynaProperty;
import org.junit.Test;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;
import org.jxls.entity.TestEmployee;
import org.jxls.util.JxlsHelper;

/**
 * This test class checks whether grouping works with DynaBeans. (Issue 182)
 */
public class DynaBeanTest {

    /**
     * This testcase tests grouping with DynaBean. (Fixed with issue 182)
     * It also checks whether DynaBeans work without grouping. (Worked before because of JEXL)
     */
    @Test
    public void groupingWithDynaBean() throws Exception {
        // Prepare
        Context context = new Context();
        context.putVar("employees", generateDynaSampleEmployeeData());
        String out = "target/DynaBeanTest_output.xlsx";

        // Test
        try (InputStream is = getClass().getResourceAsStream("DynaBeanTest.xlsx")) {
            try (OutputStream os = new FileOutputStream(out)) {
                JxlsHelper.getInstance().processTemplate(is, os, context);
            }
        }
        
        // Verify
        try (TestWorkbook xls = new TestWorkbook(new File(out))) {
            xls.selectSheet("grouping");
            assertEquals("Elsa", xls.getCellValueAsString(2, 1));
            assertEquals("John", xls.getCellValueAsString(3, 1));
            assertEquals("Oleg", xls.getCellValueAsString(4, 1));
            
            xls.selectSheet("simple"); // no grouping
            assertEquals("Elsa", xls.getCellValueAsString(2, 1));
            assertEquals("Oleg", xls.getCellValueAsString(3, 1));
            assertEquals("John", xls.getCellValueAsString(4, 1));
        }
    }

    private List<DynaBean> generateDynaSampleEmployeeData() throws Exception {
        DynaClass dynaClass = new BasicDynaClass("Employee", null,
                new DynaProperty[] { new DynaProperty("name", String.class), });
        List<DynaBean> employeesDyna = new ArrayList<>();
    
        DynaBean elsa = dynaClass.newInstance();
        elsa.set("name", "Elsa");
        employeesDyna.add(elsa);
    
        DynaBean oleg = dynaClass.newInstance();
        oleg.set("name", "Oleg");
        employeesDyna.add(oleg);
    
        DynaBean john = dynaClass.newInstance();
        john.set("name", "John");
        employeesDyna.add(john);
    
        return employeesDyna;
    }

    /**
     * This testcase tests grouping with Java bean. It also checks whether Java beans work without grouping.
     */
    @Test
    public void groupingWithJavaBean() throws Exception {
        // Prepare
        Context context = new Context();
        context.putVar("employees", generateStaticSampleEmployeeData());
        String out = "target/DynaBeanTest_output.xlsx";

        // Test
        try (InputStream is = getClass().getResourceAsStream("DynaBeanTest.xlsx")) {
            try (OutputStream os = new FileOutputStream(out)) {
                JxlsHelper.getInstance().processTemplate(is, os, context);
            }
        }
        
        // Verify
        try (TestWorkbook xls = new TestWorkbook(new File(out))) {
            xls.selectSheet("grouping");
            assertEquals("Elsa", xls.getCellValueAsString(2, 1));
            assertEquals("John", xls.getCellValueAsString(3, 1));
            assertEquals("Oleg", xls.getCellValueAsString(4, 1));
            
            xls.selectSheet("simple"); // no grouping
            assertEquals("Elsa", xls.getCellValueAsString(2, 1));
            assertEquals("Oleg", xls.getCellValueAsString(3, 1));
            assertEquals("John", xls.getCellValueAsString(4, 1));
        }
    }

    private List<TestEmployee> generateStaticSampleEmployeeData() throws Exception {
        List<TestEmployee> employees = new ArrayList<>();
        employees.add(new TestEmployee("", "Elsa", "", "", 0));
        employees.add(new TestEmployee("", "Oleg", "", "", 0));
        employees.add(new TestEmployee("", "John", "", "", 0));
        return employees;
    }
}
