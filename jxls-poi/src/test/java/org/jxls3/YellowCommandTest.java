package org.jxls3;

import java.util.HashMap;
import java.util.Map;

import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.TestWorkbook;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.command.AbstractCommand;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.Size;
import org.jxls.entity.Employee;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;
import org.jxls.util.Util;

public class YellowCommandTest {

	@Test
	public void test() {
		// Prepare
        Map<String,Object> data = new HashMap<>();
        data.put("employees", Employee.generateSampleEmployeeData());

        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass());
        tester.test(data, JxlsPoiTemplateFillerBuilder.newInstance().withAreaBuilder(new XlsCommentAreaBuilder() {
        	// TODO add YellowCommand
        }));
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet(0);
            // XXX Baustelle
			System.out.println("C1: " + w._getCell(1, 3).getCellStyle().getFillBackgroundColorColor()); // not null ok
			System.out.println("C2: " + w._getCell(2, 3).getCellStyle().getFillBackgroundColorColor()); // null ok
			System.out.println("C6: " + w._getCell(6, 3).getCellStyle().getFillBackgroundColorColor()); // TODO must non null
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
	        	// TODO make cell yellow
	        }
			return new Size(0, 0);
		}
	}
}
