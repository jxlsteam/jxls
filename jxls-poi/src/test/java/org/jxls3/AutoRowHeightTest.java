package org.jxls3;

import static org.junit.Assert.assertTrue;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Sheet;
import org.junit.Ignore;
import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.TestWorkbook;
import org.jxls.entity.Employee;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

public class AutoRowHeightTest {

    @Test
    @Ignore("Test fails yet")
    public void testAutoRow() {
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass());

        Map<String, Object> data = createData();
        tester.test(data, JxlsPoiTemplateFillerBuilder.newInstance());

        try (TestWorkbook w = tester.getWorkbook()) {
            Sheet sheet = w.selectSheet("Employees");
            sheet.getRow(0).getHeight();

            short lenOfRecord1 = sheet.getRow(1).getHeight();
            short lenOfRecord2 = sheet.getRow(2).getHeight();
            assertTrue("Second row of data should be higher than others", lenOfRecord1 < lenOfRecord2);
        }
    }

    private Map<String, Object> createData() {
        List<Employee> employees = Employee.generateSampleEmployeeData();
        employees.get(1).setName("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua");

        Map<String, Object> data = new HashMap<>();
        data.put("employees", employees);
        return data;
    }
}
