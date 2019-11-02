package org.jxls.transform.poi;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.Test;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

import java.io.*;
import java.util.Arrays;
import java.util.List;

import static org.junit.Assert.assertEquals;

public class ConditionalFormattingTest {

    private static final double EPSILON = 0.001;

    @Test
    public void shouldCopyConditionalFormatInEachCommandLoop() throws IOException {
        InputStream is = ConditionalFormattingTest.class.getResourceAsStream("cond_format_template.xlsx");
        String outputFileName = "target/cond_format_output.xlsx";
        File outputFile = new File(outputFileName);
        FileOutputStream out = new FileOutputStream(outputFile);
        Context context = new Context();
        List<Integer> list = Arrays.asList(2, 1, 4, 3, 5);
        context.putVar("numbers", list);
        context.putVar("val1", 0);
        context.putVar("val2", 7);
        JxlsHelper.getInstance().processTemplate(is, out, context);
        Workbook workbook = WorkbookFactory.create(outputFile);
        for(int i = 0; i < list.size(); i++){
            assertEquals(list.get(i).doubleValue(), PoiTestHelper.cellNumericValue(workbook, i + 1, 1), EPSILON);
        }
        Sheet sheet = workbook.getSheetAt(0);
        SheetConditionalFormatting sheetConditionalFormatting = sheet.getSheetConditionalFormatting();
        int conditionalFormattingCount = 0;
        for( int i = 0; i < sheetConditionalFormatting.getNumConditionalFormattings(); i++){
            ConditionalFormatting conditionalFormatting = sheetConditionalFormatting.getConditionalFormattingAt(i);
            CellRangeAddress[] ranges = conditionalFormatting.getFormattingRanges();
            if (ranges.length > 0){
                conditionalFormattingCount++;
            }
        }
        assertEquals(list.size() + 2, conditionalFormattingCount);
    }
}
