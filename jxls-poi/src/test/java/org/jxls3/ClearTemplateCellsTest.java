package org.jxls3;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

import org.junit.Assert;
import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.TestWorkbook;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

public class ClearTemplateCellsTest {

	@Test
	public void unusedExpressionsStay() {
		check(false, "${e.name}");
	}

	@Test
	public void unusedExpressionsAreCleared() {
		check(true, "");
	}

	private void check(boolean clearTemplateCells, String expectation) {
		// Prepare
		Map<String,Object> data = new HashMap<>();
		data.put("employees", new ArrayList<>());
		
		// Test
		Jxls3Tester tester = Jxls3Tester.xlsx(getClass());
		tester.test(data, JxlsPoiTemplateFillerBuilder.newInstance().withClearTemplateCells(clearTemplateCells));

		// Verify
		try (TestWorkbook w = tester.getWorkbook()) {
			w.selectSheet(0);
			Assert.assertEquals(expectation, w.getCellValueAsString(2, 1));
		}
	}
}
