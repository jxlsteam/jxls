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
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.TestWorkbook;
import org.jxls.area.Area;
import org.jxls.area.XlsArea;
import org.jxls.builder.JxlsOutputFile;
import org.jxls.builder.JxlsStreaming;
import org.jxls.builder.JxlsTemplateFiller;
import org.jxls.command.EachCommand;
import org.jxls.command.IfCommand;
import org.jxls.entity.Employee;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

public class IfTest {

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
        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass());
        tester.test(data(), JxlsPoiTemplateFillerBuilder.newInstance().withStreaming(streaming));
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet("Employees");
            assertEquals("Elsa", w.getCellValueAsString(2, 1)); // A2 
            assertEquals("1969-05-30T00:00", w.getCellValueAsLocalDateTime(6, 2).toString()); // B6 
            assertEquals(2500d, w.getCellValueAsDouble(4, 3), 0.005d); // C4
            assertNull(w.getCellValueAsString(7, 1)); // A7
        }
    }

    @Test
    public void noElse() {
        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass(), "noElse");
        tester.test(data(), JxlsPoiTemplateFillerBuilder.newInstance());
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet("Employees");
            assertEquals("Elsa", w.getCellValueAsString(2, 1)); 
            assertEquals("Maria", w.getCellValueAsString(3, 1)); 
            assertNull(w.getCellValueAsString(4, 1)); 
        }
    }

    @Test
    public void noComments() { // TODO better name: noNotes
        // Test
        InputStream template = getClass().getResourceAsStream("IfTest_noComments.xlsx");
        File out = new File("target/IfTest_noComments_output.xlsx");
        out.getParentFile().mkdirs();
        JxlsTemplateFiller builder = JxlsPoiTemplateFillerBuilder.newInstance().withAreaBuilder((transformer, ctc) -> {
        	List<Area> areas = new ArrayList<>();	
        	XlsArea area = new XlsArea("Employees!A1:C2", transformer);
        	areas.add(area);
        	XlsArea eachArea = new XlsArea("Employees!A2:C2", transformer);
        	area.addCommand("A2:C2", new EachCommand("e", "employees", eachArea));
        	eachArea.addCommand("A2:C2", new IfCommand("e.payment<2000", new XlsArea("Employees!A2:C2", transformer), new XlsArea("Employees!A3:C3", transformer)));
        	return areas;
        }).withTemplate(template).build();
		builder.fill(data(), new JxlsOutputFile(out));

        // Verify
        try (TestWorkbook w = new TestWorkbook(out)) {
            w.selectSheet("Employees");
            assertEquals("Elsa", w.getCellValueAsString(2, 1)); 
            assertEquals("Oleg", w.getCellValueAsString(3, 1)); 
            assertEquals("Neil", w.getCellValueAsString(4, 1)); 
            assertEquals("Maria", w.getCellValueAsString(5, 1)); 
            assertEquals("John", w.getCellValueAsString(6, 1)); 
        }
    }

	private Map<String, Object> data() {
		Map<String,Object> data = new HashMap<>();
        data.put("employees", Employee.generateSampleEmployeeData());
		return data;
	}
}
