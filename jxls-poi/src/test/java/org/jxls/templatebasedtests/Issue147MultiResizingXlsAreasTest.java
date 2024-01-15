package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;

import java.io.IOException;
import java.io.InputStream;
import java.util.Map;

import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.TestWorkbook;
import org.jxls.area.Area;
import org.jxls.builder.JxlsOptions;
import org.jxls.builder.JxlsTemplateFiller;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.ContextImpl;
import org.jxls.common.Size;
import org.jxls.entity.Employee;
import org.jxls.formula.StandardFormulaProcessor;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

/**
 * Issue 147
 *
 * In a custom processTemplate() that supports multiple contiguous XlsAreas, row heights are not copied from the
 * source cell to the destination properly. Multiple contiguous XlsAreas occur in JXLS Markup templates where the
 * target sheet is the template sheet. The XlsArea is applied at a target that was shifted as a result of the prior
 * XlsArea processing (i.e. area grew to ForEach command), and XlsArea.updateRowHeights is using an offset from
 * target cell start, rather than original area start.
 */
public class Issue147MultiResizingXlsAreasTest {
//    private final AreaBuilder areaBuilder = new XlsCommentAreaBuilder();
    private Context context;
    
    @Test
    public void test() throws IOException {
        // Prepare
        context = new ContextImpl();
        context.putVar("employees", Employee.generateSampleEmployeeData());
        
        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass());
        // install own MyJxlsTemplateFiller
        JxlsPoiTemplateFillerBuilder builder = new JxlsPoiTemplateFillerBuilder() {
            @Override
            public JxlsTemplateFiller build() {
                return new MyJxlsTemplateFiller(getOptions(), template);
            }
        };
        tester.test(context.toMap(), builder
                .withExceptionThrower()
                .withTemplate(getClass().getResourceAsStream("Issue147MultiResizingXlsAreasTest.xlsx")));
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet("Template");
            assertEquals(50 * 15, w.getRowHeight(1));  // First area header
            assertEquals(20 * 15, w.getRowHeight(3));  // First table header
            assertEquals(30 * 15, w.getRowHeight(4));  // First table row
            assertEquals(50 * 15, w.getRowHeight(11)); // Second area header
            assertEquals(20 * 15, w.getRowHeight(13)); // Second table header
            assertEquals(30 * 15, w.getRowHeight(14)); // Second table row
            assertEquals(38 * 15, w.getRowHeight(21)); // Footer
        }
    }
    
    public class MyJxlsTemplateFiller extends JxlsTemplateFiller {

        protected MyJxlsTemplateFiller(JxlsOptions options, InputStream template) {
            super(options, template);
        }
        
        @Override
        protected void processAreas(Map<String, Object> data) {
            areas = options.getAreaBuilder().build(transformer, true);
            Size delta = Size.ZERO_SIZE;
            for (Area area : areas) {
                CellRef targetCellRef = new CellRef(area.getStartCellRef().getSheetName(),
                        area.getStartCellRef().getRow() + delta.getHeight(),
                        area.getStartCellRef().getCol() + delta.getWidth());
                Size startSize = area.getSize();
                Size endSize = area.applyAt(targetCellRef, context);
                delta = delta.add(endSize.minus(startSize));
                area.processFormulas(new StandardFormulaProcessor());
            }
        }
    }
}
