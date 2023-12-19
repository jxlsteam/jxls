package org.jxls3;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNull;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.TestWorkbook;
import org.jxls.area.Area;
import org.jxls.area.XlsArea;
import org.jxls.builder.JxlsOutputFile;
import org.jxls.builder.JxlsTemplateFiller;
import org.jxls.command.GridCommand;
import org.jxls.entity.Employee;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

public class GridTest {

    @Test
    public void test() {
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass());
        tester.test(data(), JxlsPoiTemplateFillerBuilder.newInstance());
        
        try (TestWorkbook w = tester.getWorkbook()) {
            verify(w);
        }
    }
    
    @Test
    public void noComments() {
        InputStream template = getClass().getResourceAsStream("GridTest_noComments.xlsx");
        File out = new File("target/GridTest_noComments_output.xlsx");
        out.getParentFile().mkdirs();
        JxlsTemplateFiller builder = JxlsPoiTemplateFillerBuilder.newInstance().withAreaBuilder((transformer, ctc) -> {
            List<Area> areas = new ArrayList<>();   
            XlsArea area = new XlsArea("Sheet1!A1:A4", transformer);
            areas.add(area);
            GridCommand grid = new GridCommand("headers", "data", new XlsArea("Sheet1!A3:A3", transformer), new XlsArea("Sheet1!A4:A4", transformer));
            grid.setFormatCells("BigDecimal:C1,Date:D1");
            area.addCommand("A3:A4", grid);
            return areas;
        }).withTemplate(template).build();
        builder.fill(data(), new JxlsOutputFile(out));

        try (TestWorkbook w = new TestWorkbook(out)) {
            verify(w);
        }
    }
    
    /**
     * Allow Excel formulas to work with jx:grid
     */
    @Test
    public void issueB090() throws IOException {
        Map<String, Object> data = data();

        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass(), "issueB090");
        tester.test(data, JxlsPoiTemplateFillerBuilder.newInstance());
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet(0);
            assertEquals(1, w.getCellValueAsDouble(2, 2), 0.1d);
        }
    }

    private Map<String, Object> data() {
        Map<String, Object> data = new HashMap<>();
        data.put("headers", Arrays.asList("Name", "Birthday", "Payment"));
        data.put("data", createGridData(Employee.generateSampleEmployeeData()));
        return data;
    }

    private List<List<Object>> createGridData(List<Employee> employees) {
        List<List<Object>> data = new ArrayList<>();
        for (Employee employee : employees) {
            data.add(convertEmployeeToList(employee));
        }
        return data;
    }

    private List<Object> convertEmployeeToList(Employee employee) {
        List<Object> list = new ArrayList<>();
        list.add(employee.getName());
        list.add(employee.getBirthDate());
        list.add(employee.getPayment());
        return list;
    }

    private void verify(TestWorkbook w) {
        w.selectSheet(0);
        assertEquals("Name", w.getCellValueAsString(3, 1)); 
        assertEquals("Birthday", w.getCellValueAsString(3, 2)); 
        assertEquals("Payment", w.getCellValueAsString(3, 3)); 
        assertEquals("Elsa", w.getCellValueAsString(4, 1)); // A4 
        assertEquals("1975-10-05T00:00", w.getCellValueAsLocalDateTime(6, 2).toString()); // B6 
        assertEquals(1500d, w.getCellValueAsDouble(4, 3), 0.005d); // C4
        assertNull(w.getCellValueAsString(9, 1)); // A9
    }
}
