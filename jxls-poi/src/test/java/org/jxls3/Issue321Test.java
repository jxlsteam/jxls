package org.jxls3;

import static org.junit.Assert.assertEquals;

import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.junit.Before;
import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.TestWorkbook;
import org.jxls.entity.Employee;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

public class Issue321Test {

    int dataSize;
    Map<String, Object> data;

    @Before
    public void setUp() {
        List<Employee> list = Employee.generateSampleEmployeeData();
        data = new HashMap<>();
        data.put("employees", list);
        data.put("employees2", list);
        dataSize = list.size();
    }


    @Test
    public void testMoreCommandsBeforeAThirdCommand() throws IOException {
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass());
        tester.test(data, JxlsPoiTemplateFillerBuilder.newInstance());
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet(0);
            int afterData = dataSize + 2 /* header rows */ + 1 /* 1-based */;
            assertEquals("d", w.getCellValueAsString(afterData, 2));
            assertEquals("LAST ONE", w.getCellValueAsString(afterData + 3, 2));
        }
    }

}
