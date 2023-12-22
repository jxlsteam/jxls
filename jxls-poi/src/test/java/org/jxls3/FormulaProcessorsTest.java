package org.jxls3;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNull;

import java.util.HashMap;
import java.util.Map;

import org.junit.Assert;
import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.TestWorkbook;
import org.jxls.entity.Employee;
import org.jxls.formula.FastFormulaProcessor;
import org.jxls.formula.StandardFormulaProcessor;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

public class FormulaProcessorsTest {

	@Test
	public void withFastFormulaProcessor() {
		Assert.assertTrue(JxlsPoiTemplateFillerBuilder.newInstance().withFastFormulaProcessor() //
				.getFormulaProcessor().getClass() == FastFormulaProcessor.class);
	}

	@Test
	public void createFastFormulaProcessor() {
		Assert.assertTrue(JxlsPoiTemplateFillerBuilder.newInstance().withFormulaProcessor(new FastFormulaProcessor()) //
				.getFormulaProcessor().getClass() == FastFormulaProcessor.class);
	}

	@Test
	public void defaultFormulaProcessor() {
		Assert.assertTrue(JxlsPoiTemplateFillerBuilder.newInstance() //
				.getFormulaProcessor().getClass() == StandardFormulaProcessor.class);
	}

	@Test
	public void none() {
		// Prepare
        Map<String,Object> data = new HashMap<>();
        data.put("employees", Employee.generateSampleEmployeeData());

        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass(), "none");
		JxlsPoiTemplateFillerBuilder builder = JxlsPoiTemplateFillerBuilder.newInstance().withFormulaProcessor(null);
		tester.test(data, builder);
		
		// Verify
        try (TestWorkbook w = tester.getWorkbook()) {
		    w.selectSheet(0);
			for (int row = 2; row <= 6; row++) {
				// Rows 3 to 6 have the value of row 2 because the FormulaProcessor did not change the formulas for each row.
				assertEquals(3000d, w.getCellValueAsDouble(row, 5), 0.005d);
			}
		}
        assertNull(builder.getFormulaProcessor());
	}
}
