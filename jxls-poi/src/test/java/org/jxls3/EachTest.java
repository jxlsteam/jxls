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
}
