package org.jxls3;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNull;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.TestWorkbook;
import org.jxls.command.AbstractCommand;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.Size;
import org.jxls.entity.Employee;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;
import org.jxls.transform.poi.PoiTransformer;
import org.jxls.util.Util;

public class YellowCommandTest {

	@Test
	public void test() {
		// Prepare
        Map<String,Object> data = new HashMap<>();
        data.put("employees", Employee.generateSampleEmployeeData());

        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass());
		tester.test(data, JxlsPoiTemplateFillerBuilder.newInstance().withCommand("yellow", YellowCommand.class));
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet(0);
            assertEquals("yellow", w.getCellValueAsString(6, 3));
            assertNull(w.getCellValueAsString(5, 3));
        }
	}
	
	/** Makes 1 cell yellow */
	public static class YellowCommand extends AbstractCommand {
		private String condition;

		@Override
		public String getName() {
			return "yellow";
		}

		public String getCondition() {
			return condition;
		}

		public void setCondition(String condition) {
			this.condition = condition;
		}

		@Override
		public Size applyAt(CellRef cellRef, Context context) {
	        Boolean conditionResult = Util.isConditionTrue(getTransformationConfig().getExpressionEvaluator(), condition, context);
	        if (conditionResult.booleanValue()) {
	    		Row row = ((PoiTransformer) getTransformer()).getWorkbook().getSheet(cellRef.getSheetName()).getRow(cellRef.getRow());
	    		Cell cell = row.getCell(cellRef.getCol());
	    		if (cell == null) {
	    			cell = row.createCell(cellRef.getCol());
	    		}
	    		cell.setCellValue("yellow");
	        }
			return new Size(1, 1);
		}
	}
}
