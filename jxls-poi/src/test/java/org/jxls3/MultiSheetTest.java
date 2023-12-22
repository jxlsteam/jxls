package org.jxls3;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.junit.Assert;
import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.TestWorkbook;
import org.jxls.builder.KeepTemplateSheet;
import org.jxls.entity.Employee;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

public class MultiSheetTest {
	
	@Test
	public void multisheet() {
        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass());
        tester.test(data(), JxlsPoiTemplateFillerBuilder.newInstance().withKeepTemplateSheet(KeepTemplateSheet.KEEP));
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            try {
				w.selectSheet("notexisting");
				Assert.fail("failed: check if non-existing sheets cause an error");
			} catch (IllegalArgumentException expected) {
			}
            w.selectSheet("template");
            // check if all sheets have been created and filled:
            verifySheet(w, "Elsa");
            verifySheet(w, "Oleg");
            verifySheet(w, "Neil");
            verifySheet(w, "Maria");
            verifySheet(w, "John");
            assertEquals("1969-05-30T00:00", w.getCellValueAsLocalDateTime(3, 2).toString()); 
        }
	}

	@Test
	public void deleteTemplateSheet() {
        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass());
        tester.test(data(), JxlsPoiTemplateFillerBuilder.newInstance());
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            verifySheet(w, "Elsa");
            verifySheet(w, "John");
            try {
            	// check if sheet "template" has been deleted
				w.selectSheet("template");
				Assert.fail("Sheet 'template' has not been deleted!");
			} catch (IllegalArgumentException expected) {
				assertTrue(expected.getMessage().contains("exist"));
				assertFalse(expected.getMessage().contains("visible"));
			}
        }
    }

	@Test
	public void hideTemplateSheet() {
        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass());
        tester.test(data(), JxlsPoiTemplateFillerBuilder.newInstance().withKeepTemplateSheet(KeepTemplateSheet.HIDE));
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            verifySheet(w, "Elsa");
            verifySheet(w, "John");
            try {
            	// check if sheet "template" has been hidden
				w.selectSheet("template");
				Assert.fail("Sheet 'template' has not been hidden!");
			} catch (IllegalArgumentException expected) {
				assertTrue(expected.getMessage().contains("visible"));
			}
        }
    }

	@Test
	public void getter() {
		// Prepare
		Map<String, Object> data = new HashMap<>();
        data.put("employees", Employee.generateSampleEmployeeData());

        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass(), "getter");
        tester.test(data(), JxlsPoiTemplateFillerBuilder.newInstance());
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            verifySheet(w, "Elsa");
            verifySheet(w, "Oleg");
            verifySheet(w, "Neil");
            verifySheet(w, "Maria");
            verifySheet(w, "John");
        }
	}

	private Map<String, Object> data() {
		Map<String, Object> data = new HashMap<>();
        List<Employee> employees = Employee.generateSampleEmployeeData();
		data.put("employees", employees);
        data.put("sheetNames", employees.stream().map(i -> i.getName()).toList());
		return data;
	}

	private void verifySheet(TestWorkbook w, String name) {
        w.selectSheet(name);
        assertEquals(name, w.getCellValueAsString(2, 2)); 
	}
}
