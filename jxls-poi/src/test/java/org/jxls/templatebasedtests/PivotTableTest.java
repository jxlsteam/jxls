package org.jxls.templatebasedtests;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;

import org.junit.Assert;
import org.junit.Test;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;
import org.jxls.entity.Employee;
import org.jxls.util.JxlsHelper;
import org.jxls.util.JxlsNationalLanguageSupport;

public class PivotTableTest {

    /**
     * Issue B155: Pivot table does not work with NLS
     */
    @Test
    public void nls() throws Exception {
        // Prepare
        final Context context = new Context();
        context.putVar("employees", getTestData());
        final Properties resourceBundle = new Properties();
        resourceBundle.put("name", "Name (EN)");
        resourceBundle.put("salary", "Salary (EN)");
        
        // Test
        JxlsNationalLanguageSupport nls = new JxlsNationalLanguageSupport() {
            @Override
            protected String translate(String name, String fallback) {
                return resourceBundle.getProperty(name, fallback);
            }
        };
        File temp = nls.process(getClass().getResourceAsStream(getClass().getSimpleName() + ".xlsx")); // do preprocessing of template file
        File out = new File("target/" + getClass().getSimpleName() + "_output.xlsx");
        try (InputStream is = new FileInputStream(temp)) {
            try (OutputStream os = new FileOutputStream(out)) {
                JxlsHelper.getInstance().processTemplate(is, os, context);
            }
        }
        temp.delete();
        
        // Verify
        try (TestWorkbook w = new TestWorkbook(out)) {
            w.selectSheet("Employees");
            Assert.assertEquals("Name (EN)", w.getCellValueAsString(1, 1));
            Assert.assertEquals("BU", w.getCellValueAsString(1, 2));
            Assert.assertEquals("Salary (EN)", w.getCellValueAsString(1, 3));
            Assert.assertEquals("Sven", w.getCellValueAsString(2, 1));
            w.selectSheet("Crosstab");
            Assert.assertTrue(w.getCellValueAsString(7, 4).contains("Salary (EN)"));
            // It's not possible to verify the PivotTable values because it's calculated when opening in Excel.
            // Best verification is to look at the result file using MS Excel.
        }
    }

    private List<Employee> getTestData() {
        List<Employee> list = new ArrayList<>();
        add(list, "Sven", "Mayor", 100000);
        add(list, "Christiane", "Finance", 30000);
        add(list, "John", "Main", 50000);
        add(list, "Betty", "Finance", 45000);
        add(list, "Waldemar", "Main", 60000);
        return list;
    }

    private void add(List<Employee> list, String name, String department, double salary) {
        Employee e = new Employee(name, null, salary, 0);
        e.setBuGroup(department);
        list.add(e);
    }
}
