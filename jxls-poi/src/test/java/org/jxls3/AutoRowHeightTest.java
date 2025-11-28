package org.jxls3;

import static org.junit.Assert.assertTrue;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.junit.Ignore;
import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.TestWorkbook;
import org.jxls.entity.Employee;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

public class AutoRowHeightTest {

    @Test
    public void testAutoRow() {
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass());

        Map<String, Object> data = createData();
        tester.test(data, JxlsPoiTemplateFillerBuilder.newInstance());

        try (TestWorkbook w = tester.getWorkbook()) {
            Sheet sheet = w.selectSheet("Employees");
            XSSFRow row1 = (XSSFRow) sheet.getRow(1);
            XSSFRow row2 = (XSSFRow) sheet.getRow(2);

            // this test should be improved.
            // We are not able checking row height, so we just check that height is not explicitly set.
            assertTrue("Height should not be set", !row1.getCTRow().isSetHt());
            assertTrue("Height should not be set", !row2.getCTRow().isSetHt());
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
