package org.jxls.templatebasedtests;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Assert;
import org.junit.Test;
import org.jxls.TestWorkbook;
import org.jxls.area.Area;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.transform.poi.PoiContext;
import org.jxls.transform.poi.PoiTransformer;

/**
 * Test for issue 153
 * 
 * Issue in Excel Output while using SXSSF Transformer with JXLS >= 2.7.0
 * cause: commit 5354beaf
 */
public class IssueSxssfTransformerTest {

    @Test
    public void test() throws IOException {
        // Test
        final String output = "target/" + getClass().getSimpleName() + "_output.xlsx";
        try (InputStream is = getClass().getResourceAsStream(getClass().getSimpleName() + ".xlsx")) {
            try (OutputStream os = new FileOutputStream(output)) {
                final Workbook workbook = WorkbookFactory.create(is);
                final int activeSheetIndex = workbook.getActiveSheetIndex();
                PoiTransformer transformer = PoiTransformer.createSxssfTransformer(workbook, 1000, true);
                final List<Area> excelAreas = new XlsCommentAreaBuilder(transformer).build();
                for (final Area excelArea : excelAreas) {
                    processArea(excelArea);
                }
                workbook.setForceFormulaRecalculation(true);
                workbook.setActiveSheet(activeSheetIndex);
                SXSSFWorkbook workbook2 = (SXSSFWorkbook) transformer.getWorkbook();
                workbook2.write(os);
            }
        }
        
        // Verify
        try (TestWorkbook w = new TestWorkbook(new File(output))) {
            w.selectSheet(1);
            Assert.assertEquals("Manager:", w.getCellValueAsString(3, 1)); // A3
        }
    }

    private void processArea(Area area) {
        Context context = prepareContext();
        final CellRef ref = new CellRef("Result", 0, 0);
        area.applyAt(ref, context);
    }

    private Context prepareContext() {
        final Context context = new PoiContext();
        context.getConfig().setIsFormulaProcessingRequired(false);

        ArrayList<Map<String,String>> mapArrayList = new ArrayList<>();
        mapArrayList.add(Collections.singletonMap("entity", "ABC"));
        mapArrayList.add(Collections.singletonMap("entity", "BDE"));
        mapArrayList.add(Collections.singletonMap("entity", "EFG"));

        ArrayList<Map<String,String>> mapOrgArrayList = new ArrayList<>();
        mapOrgArrayList.add(Collections.singletonMap("entity", "ABC"));
        mapOrgArrayList.add(Collections.singletonMap("entity", "BDE"));
        mapOrgArrayList.add(Collections.singletonMap("entity", "EFG"));

        context.putVar("departmentsName", mapArrayList);
        context.putVar("departmentsOrgName", mapOrgArrayList);
        return context;
    }
}
