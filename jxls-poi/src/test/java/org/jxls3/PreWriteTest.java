package org.jxls3;

import java.util.HashMap;
import java.util.Map;

import org.junit.Assert;
import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.TestWorkbook;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

public class PreWriteTest {

	@Test
    public void recalculateFormulasBeforeSaving() {
		// Prepare
        Map<String, Object> data = data();

        // Test with default value true
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass(), "recalculateFormulasBeforeSaving");
        JxlsPoiTemplateFillerBuilder builder = JxlsPoiTemplateFillerBuilder.newInstance();
		tester.test(data, builder);
        verify115_9(tester);
        
        // Test with false
		builder.withRecalculateFormulasBeforeSaving(false);
		tester.test(data, builder);
        verify0(tester);

        // Test with true
		builder.withRecalculateFormulasBeforeSaving(true);
		tester.test(data, builder);
        verify115_9(tester);
    }

	@Test
    public void recalculateFormulasOnOpening() {
        // Prepare
        Map<String, Object> data = data();

        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass(), "recalculateFormulasBeforeSaving");
        JxlsPoiTemplateFillerBuilder builder = JxlsPoiTemplateFillerBuilder.newInstance()
        		.withRecalculateFormulasBeforeSaving(false).withRecalculateFormulasOnOpening(true);
		tester.test(data, builder);
		verify0(tester); // Reading the value of B3 using POI results in 0.
		// You can only verify this manually using Excel for opening the result file: value of B3 must be 115.9
	}

	private Map<String, Object> data() {
		Map<String,Object> data = new HashMap<>();
        data.put("a", Double.valueOf(12.3));
        data.put("b", Double.valueOf(103.6));
		return data;
	}

	private void verify0(Jxls3Tester tester) {
		try (TestWorkbook w = tester.getWorkbook()) {
	        w.selectSheet(0);
	        Assert.assertEquals(Double.valueOf(0d), w.getCellValueAsDouble(3, 2), 0.005d);
	    }
	}

	private void verify115_9(Jxls3Tester tester) {
		try (TestWorkbook w = tester.getWorkbook()) {
	        w.selectSheet(0);
	        Assert.assertEquals(Double.valueOf(115.9), w.getCellValueAsDouble(3, 2), 0.005d);
	    }
	}
}
