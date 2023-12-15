package org.jxls3;

import static org.junit.Assert.assertTrue;

import java.util.HashMap;
import java.util.Map;

import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.command.EachCommand;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.Size;
import org.jxls.entity.Employee;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

/**
 * Extend EachCommand with subtotal functionality for showing how to extend an existing Jxls command.
 */
public class SubtotalTest {
	private static boolean subtotalActionCalled;
	
	@Test
	public void test() {
		// Prepare
        Map<String,Object> data = new HashMap<>();
        data.put("employees", Employee.generateSampleEmployeeData());
        subtotalActionCalled = false;

        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass());
		tester.test(data, JxlsPoiTemplateFillerBuilder.newInstance().withCommand(EachCommand.COMMAND_NAME, EachCommandWithSubtotal.class));
        
        // Verify
		assertTrue("subtotal action has not been called!", subtotalActionCalled);
	}
	
	public static class EachCommandWithSubtotal extends EachCommand {
		private String subtotal;
		
		public String getSubtotal() {
			return subtotal;
		}

		public void setSubtotal(String subtotal) {
			this.subtotal = subtotal;
		}

		@Override
		public Size applyAt(CellRef cellRef, Context context) {
			if (subtotal != null && !subtotal.isBlank()) {
				// subtotal logic ...
				subtotalActionCalled = true;
			}
			return super.applyAt(cellRef, context);
		}
	}
}
