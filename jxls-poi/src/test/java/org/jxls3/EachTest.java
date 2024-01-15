package org.jxls3;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;
import static org.jxls.builder.JxlsStreaming.AUTO_DETECT;
import static org.jxls.builder.JxlsStreaming.STREAMING_OFF;
import static org.jxls.builder.JxlsStreaming.STREAMING_ON;
import static org.jxls.builder.JxlsStreaming.streamingWithGivenSheets;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

import org.junit.Ignore;
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
        Map<String, Object> data = new HashMap<>();
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

	@Ignore("does not work on Github Maven")
	@Test(expected = JxlsTemplateFillException.class)
	public void testOutputCannotBeWritten() {
        // Prepare
        Map<String,Object> data = new HashMap<>();
        data.put("employees", Employee.generateSampleEmployeeData());
        InputStream template = getClass().getResourceAsStream(getClass().getSimpleName() + ".xlsx");
        assertNotNull(template);
        JxlsPoiTemplateFillerBuilder builder = JxlsPoiTemplateFillerBuilder.newInstance().withTemplate(template);

        // Test
        builder.buildAndFill(data, new File(":/"/*error*/));
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

	// This is also a direction=RIGHT test.
    @Test
    public void ignoreProps() throws ParseException {
    	// Prepare
		List<EmployeeWithDepartments> employees = new ArrayList<>();
		employees.add(new EmployeeWithDepartments("Kylie", "1970-Jul-10", 2100, "Main"));
		employees.add(new EmployeeWithDepartments("Mariah", "1943-Jan-01", 0));
		employees.add(new EmployeeWithDepartments("Sascha", "1965-Mar-19", 600, "D", "E", "F", "G", "H"));
		Map<String, Object> data = new HashMap<>();
        data.put("employees", employees);
        
        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass(), "ignoreProps");
        tester.test(data, JxlsPoiTemplateFillerBuilder.newInstance()
        		.withIgnoreColumnProps(true).withIgnoreRowProps(true)); // Jxls doesn't change row heights and column widths.
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet("Employees");
            assertEquals("Sascha", w.getCellValueAsString(4, 1));
            assertEquals("direction=RIGHT does not work", "H", w.getCellValueAsString(4, 8));
            assertNotEquals(w.getColumnWidth(4), w.getColumnWidth(5));
            assertNotEquals(w.getColumnWidth(5), w.getColumnWidth(6));
            assertNotEquals(w.getRowHeight(2), w.getRowHeight(3));
        }
        
        // Test
        tester = Jxls3Tester.xlsx(getClass(), "ignoreProps");
        tester.test(data, JxlsPoiTemplateFillerBuilder.newInstance()
        		.withIgnoreColumnProps(false).withIgnoreRowProps(false)); // Jxls creates rows and columns with the same width/height. (default)
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet("Employees");
            assertTrue(w.getColumnWidth(4) == w.getColumnWidth(5));
            assertTrue(w.getColumnWidth(5) == w.getColumnWidth(6));
            assertTrue(w.getRowHeight(2) == w.getRowHeight(3));
        }
    }
    
    public static class EmployeeWithDepartments extends Employee {
        private static final SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MMM-dd", Locale.US);
    	private final List<String> departments = new ArrayList<>();
        
    	public EmployeeWithDepartments(String name, String birthDate, double payment, String ...departments) throws ParseException {
        	super(name, dateFormat.parse(birthDate), payment, 0d);
    		this.departments.addAll(Arrays.asList(departments));
        }

		public List<String> getDepartments() {
			return departments;
		}
    }
    
    @Test
    public void varIndex() throws ParseException {
        // Prepare
        Map<String,Object> data = new HashMap<>();
        data.put("employees", Employee.generateSampleEmployeeData());

        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass(), "varIndex");
        tester.test(data, JxlsPoiTemplateFillerBuilder.newInstance());
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet("Employees");
            for (int i = 1; i <= 5; i++) {
                assertEquals(i, w.getCellValueAsDouble(1 + i, 3), 0.5d);
            }
        }
    }

    @Ignore("Record support not implemented")
    @Test
    public void testRecord() throws ParseException {
        // Prepare
        Map<String, Object> data = new HashMap<>();
        List<EmployeeRecord> employees = new ArrayList<>();
        employees.add(new EmployeeRecord("Walter", new SimpleDateFormat("yyyy-MMM-dd", Locale.US).parse("1970-Jul-10"), 3600d));
        data.put("employees", employees);

        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass());
        tester.test(data, JxlsPoiTemplateFillerBuilder.newInstance());
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet("Employees");
            assertEquals("Walter", w.getCellValueAsString(2, 1)); // A2 
            assertEquals("1970-07-10T00:00", w.getCellValueAsLocalDateTime(2, 2).toString()); 
            assertEquals(3600d, w.getCellValueAsDouble(2, 3), 0.005d);
        }
    }
    
    public static record EmployeeRecord(String name, Date birthDate, double payment) {
    }
}
