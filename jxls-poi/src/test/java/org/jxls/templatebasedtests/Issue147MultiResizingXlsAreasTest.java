package org.jxls.templatebasedtests;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.area.Area;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.Size;
import org.jxls.entity.Employee;
import org.jxls.transform.Transformer;
import org.jxls.util.JxlsHelper;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import static org.junit.Assert.assertEquals;
import static org.jxls.transform.poi.PoiTransformer.createTransformer;

/**
 * Issue xxx
 *
 * In a custom processTemplate that supports multiple contiguous XlsAreas, row heights are not copied from the
 * source cell to the destination properly. Multiple contiguous XlsAreas occur in JXLS Markup templates where the
 * target sheet is the template sheet. The XlsArea is applied at a target that was shifted as a result of the prior
 * XlsArea processing (i.e. area grew to ForEach command), and XlsArea.updateRowHeights is using an offset from
 * target cell start, rather than original area start
 **/
public class Issue147MultiResizingXlsAreasTest {
    private AreaBuilder areaBuilder = new XlsCommentAreaBuilder();

    public void processTemplate(Context context, Transformer transformer) throws IOException {
        Size startSize, endSize, delta = Size.ZERO_SIZE;
        areaBuilder.setTransformer(transformer);
        List<Area> xlsAreaList = areaBuilder.build();
        for (Area xlsArea : xlsAreaList) {
            CellRef targetCellRef = new CellRef( xlsArea.getStartCellRef().getSheetName(), xlsArea.getStartCellRef().getRow() + delta.getHeight(), xlsArea.getStartCellRef().getCol() + delta.getWidth());
            startSize = xlsArea.getSize();
            endSize = xlsArea.applyAt(targetCellRef, context);
            delta = delta.add(endSize.minus(startSize));
            xlsArea.processFormulas();
        }
        transformer.write();
    }

    @Test
    public void test() {
        // Prepare
        List<Employee> employees = Employee.generateSampleEmployeeData();
        Context context = new Context();
        context.putVar("employees", employees);
        
        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass());

        try (InputStream is = tester.openInputStream()) {
            try (OutputStream os = tester.openOutputStream()) {
                JxlsHelper jxls = JxlsHelper.getInstance();

                Transformer transformer = jxls.createTransformer(is, os);
                processTemplate(context, transformer);
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet("Template");
            assertEquals(50*15, w.getRowHeight(1)); //First area header
            assertEquals(20*15, w.getRowHeight(3)); //First table header
            assertEquals(30*15, w.getRowHeight(4)); //First table row
            assertEquals(50*15, w.getRowHeight(11)); //Second area header
            assertEquals(20*15, w.getRowHeight(13)); //Second table header
            assertEquals(30*15, w.getRowHeight(14)); //Second table row
            assertEquals(38*15, w.getRowHeight(21)); //Footer
       }
    }
}
