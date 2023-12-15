package org.jxls3;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNull;
import static org.jxls.builder.JxlsStreaming.AUTO_DETECT;
import static org.jxls.builder.JxlsStreaming.STREAMING_OFF;
import static org.jxls.builder.JxlsStreaming.STREAMING_ON;
import static org.jxls.builder.JxlsStreaming.streamingWithGivenSheets;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.TestWorkbook;
import org.jxls.builder.JxlsStreaming;
import org.jxls.builder.JxlsTemplateFillException;
import org.jxls.entity.Employee;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;
import org.jxls.util.CannotOpenWorkbookException;

public class EachTest {

    @Test
    public void standard() throws IOException {
        check(STREAMING_OFF);
    }

    @Test
    public void allSheets() throws IOException {
        check(STREAMING_ON);
    }

    @Test
    public void autoDetect() throws IOException {
        check(AUTO_DETECT);
    }

    @Test
    public void givenSheet() throws IOException {
        check(streamingWithGivenSheets("Employees"));
    }
    
    private void check(JxlsStreaming streaming) {
        // Prepare
        Map<String,Object> data = new HashMap<>();
        data.put("employees", Employee.generateSampleEmployeeData());

        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass());
        tester.test(data, JxlsPoiTemplateFillerBuilder.newInstance().withStreaming(streaming));
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet("Employees");
            assertEquals("Elsa", w.getCellValueAsString(2, 1)); // A2 
            assertEquals("1969-05-30T00:00", w.getCellValueAsLocalDateTime(6, 2).toString()); // B6 
            assertEquals(2500d, w.getCellValueAsDouble(4, 3), 0.005d); // C4
            assertNull(w.getCellValueAsString(7, 1)); // A7
        }
    }
    
	@Test(expected = CannotOpenWorkbookException.class)
	public void testTemplateNotFound() {
        JxlsPoiTemplateFillerBuilder.newInstance().withTemplate(getClass().getResourceAsStream("nonsense"));
	}

	@Test(expected = JxlsTemplateFillException.class)
	public void testOutputCannotBeWritten() {
        // Prepare
        Map<String,Object> data = new HashMap<>();
        data.put("employees", Employee.generateSampleEmployeeData());

        // Test
        InputStream template = getClass().getResourceAsStream(getClass().getSimpleName() + ".xlsx");
        JxlsPoiTemplateFillerBuilder.newInstance().withTemplate(template).buildAndFill(data, new File(":/"/*error*/));
	}

	/** orderBy="e.name, e.payment DESC" */
	@Test
    public void orderBy() throws IOException {
        // Prepare
        Map<String,Object> data = new HashMap<>();
        List<Employee> employees = Employee.generateSampleEmployeeData().subList(0, 4);
        employees.get(0).setName("i");
        employees.get(1).setName("Z");
        employees.get(2).setName("A");
        employees.get(3).setName("Z");
		data.put("employees", employees);

        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass(), "orderBy");
        tester.test(data, JxlsPoiTemplateFillerBuilder.newInstance());
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet("Employees");
            assertEquals("A", w.getCellValueAsString(2, 1)); 
            assertEquals("Z", w.getCellValueAsString(3, 1)); 
            assertEquals("Z", w.getCellValueAsString(4, 1)); 
            assertEquals(1700d, w.getCellValueAsDouble(4, 3), 0.005d); 
            assertEquals("i", w.getCellValueAsString(5, 1)); 
        }
    }

	@Test
    public void orderByIgnoreCase() throws IOException {
        // Prepare
        Map<String,Object> data = new HashMap<>();
        List<Employee> employees = Employee.generateSampleEmployeeData().subList(0, 3);
        employees.get(0).setName("i");
        employees.get(1).setName("Z");
        employees.get(2).setName("A");
		data.put("employees", employees);

        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass(), "orderByIgnoreCase");
        tester.test(data, JxlsPoiTemplateFillerBuilder.newInstance());
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet("Employees");
            assertEquals("A", w.getCellValueAsString(2, 1)); 
            assertEquals("i", w.getCellValueAsString(3, 1)); 
            assertEquals("Z", w.getCellValueAsString(4, 1)); 
        }
    }
}
