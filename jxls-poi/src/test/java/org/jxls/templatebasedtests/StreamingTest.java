package org.jxls.templatebasedtests;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Test;
import org.jxls.area.Area;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.transform.Transformer;
import org.jxls.transform.poi.SelectSheetsForStreamingPoiTransformer;

/**
 * Simple Streaming Demo
 */
public class StreamingTest {
    private static final int MAX = 100000;
    
    @Test
    public void test() throws IOException {
        String template = getClass().getSimpleName() + ".xlsx";
        String output = "target\\" + getClass().getSimpleName() + "_output.xlsx";
        Context context = createTestData();
        context.getConfig().setIsFormulaProcessingRequired(!false);
        Set<String> streamedSheets = new HashSet<String>();
        streamedSheets.add("Streaming");
        try (InputStream is = getClass().getResourceAsStream(template)) {
            try (OutputStream os = new FileOutputStream(output)) {
                SelectSheetsForStreamingPoiTransformer transformer = new SelectSheetsForStreamingPoiTransformer(WorkbookFactory.create(is));
                transformer.setOutputStream(os);
                transformer.setDataSheetsToUseStreaming(streamedSheets);
                transformer.setEvaluateFormulas(true);
                processTemplate(context, transformer);
                transformer.getWorkbook().write(os);
                if (transformer.getWorkbook() instanceof SXSSFWorkbook) {
                    ((SXSSFWorkbook) transformer.getWorkbook()).dispose();
                }
            }
        }
    }

    private void processTemplate(Context context, Transformer transformer) {
        XlsCommentAreaBuilder areaBuilder = new XlsCommentAreaBuilder();
        areaBuilder.setTransformer(transformer);
        List<Area> xlsAreaList = areaBuilder.build();
        for (Area xlsArea : xlsAreaList) {
            xlsArea.applyAt(new CellRef(xlsArea.getStartCellRef().getCellName()), context);
        }
    }
    
    protected Context createTestData() {
        List<TestData> list = new ArrayList<>();
        for (int i = 1; i <= MAX; i++) {
            list.add(new TestData(i));
        }
        Context ctx = new Context();
        ctx.putVar("employees", list);
        return ctx;

    }

    public static class TestData {
        private int n;
        
        public TestData(int n) {
            this.n = n;
        }
        
        public String getName() {
            return "Melanie " + n;
        }
        
        public int getAge() {
            return n;
        }
    }
}
