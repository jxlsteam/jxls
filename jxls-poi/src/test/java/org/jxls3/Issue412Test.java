package org.jxls3;

import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.TestWorkbook;
import org.jxls.entity.Employee;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

import java.util.HashMap;
import java.util.Map;

import static org.junit.Assert.assertEquals;

public class Issue412Test {

    @Test
    public void testJxlsOutputByteArray() {
        // Prepare
        Map<String, Object> data = new HashMap<>();
        data.put("employees", Employee.generateSampleEmployeeData());

        // Test
        Jxls3Tester tester = new Jxls3Tester(EachTest.class, "EachTest.xlsx");
        byte[] result = tester.testAndReturn(data, JxlsPoiTemplateFillerBuilder.newInstance());

        // Verify
        try (TestWorkbook w = tester.getWorkbook(result)) {
            w.selectSheet("Employees");
            assertEquals("Elsa", w.getCellValueAsString(2, 1)); // A2
        }
    }

}
