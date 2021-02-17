package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;

import java.util.ArrayList;
import java.util.List;

import org.apache.commons.beanutils.BasicDynaClass;
import org.apache.commons.beanutils.DynaBean;
import org.apache.commons.beanutils.DynaClass;
import org.apache.commons.beanutils.DynaProperty;
import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;
import org.jxls.entity.Employee;

/**
 * Issue B182: This test class checks whether grouping works with DynaBeans.
 */
public class DynaBeanTest {

    /**
     * This testcase tests grouping with DynaBean. (Fixed with issue B182)
     * It also checks whether DynaBeans work without grouping. (Worked before because of JEXL)
     */
    @Test
    public void groupingWithDynaBean() throws Exception {
        // Prepare
        Context context = new Context();
        context.putVar("employees", generateDynaSampleEmployeeData());

        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet("grouping");
            assertEquals("Elsa", w.getCellValueAsString(2, 1));
            assertEquals("John", w.getCellValueAsString(3, 1));
            assertEquals("Oleg", w.getCellValueAsString(4, 1));
            
            w.selectSheet("simple"); // no grouping
            assertEquals("Elsa", w.getCellValueAsString(2, 1));
            assertEquals("Oleg", w.getCellValueAsString(3, 1));
            assertEquals("John", w.getCellValueAsString(4, 1));
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

        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);
        
        // Verify
        try (TestWorkbook xls = tester.getWorkbook()) {
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

    private List<Employee> generateStaticSampleEmployeeData() throws Exception {
        List<Employee> employees = new ArrayList<>();
        employees.add(new Employee("Elsa", null, 0, 0));
        employees.add(new Employee("Oleg", null, 0, 0));
        employees.add(new Employee("John", null, 0, 0));
        return employees;
    }
}
