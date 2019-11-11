package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;

import java.util.Arrays;
import java.util.List;

import org.apache.poi.ss.usermodel.ConditionalFormatting;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;

public class ConditionalFormattingTest {
    private static final double EPSILON = 0.001;

    @Test
    public void shouldCopyConditionalFormatInEachCommandLoop() {
        // Prepare
        Context context = new Context();
        List<Integer> list = Arrays.asList(2, 1, 4, 3, 5);
        context.putVar("numbers", list);
        context.putVar("val1", 0);
        context.putVar("val2", 7);
        
        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);

        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet(0);
            for (int i = 0; i < list.size(); i++) {
                double val = w.getCellValueAsDouble(i + 2, 2);
                assertEquals(list.get(i).doubleValue(), val, EPSILON);
            }

            SheetConditionalFormatting sheetConditionalFormatting = w.getSheetConditionalFormatting();
            int conditionalFormattingCount = 0;
            for (int i = 0; i < sheetConditionalFormatting.getNumConditionalFormattings(); i++) {
                ConditionalFormatting conditionalFormatting = sheetConditionalFormatting.getConditionalFormattingAt(i);
                CellRangeAddress[] ranges = conditionalFormatting.getFormattingRanges();
                if (ranges.length > 0) {
                    conditionalFormattingCount++;
                }
            }
            assertEquals(list.size() + 2, conditionalFormattingCount);
        }
    }
}
