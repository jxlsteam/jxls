package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Map;

import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.TestWorkbook;
import org.jxls.area.Area;
import org.jxls.builder.JxlsOptions;
import org.jxls.builder.JxlsStreaming;
import org.jxls.builder.JxlsTemplateFiller;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

/**
 * Test for issue 153
 * 
 * Issue in Excel Output while using SXSSF Transformer with JXLS >= 2.7.0
 * cause: commit 5354beaf
 */
public class IssueSxssfTransformerTest {
    
    @Test
    public void test() throws IOException {
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass());
        JxlsPoiTemplateFillerBuilder builder = new JxlsPoiTemplateFillerBuilder() {
            @Override
            public JxlsTemplateFiller build() {
                return new MyJxlsTemplateFiller(getOptions(), template);
            }
        };
        tester.test(prepareContext().toMap(), builder
            .withExceptionThrower()
            .withRecalculateFormulasOnOpening(true)
            .withStreaming(JxlsStreaming.STREAMING_ON.withOptions(1000, true, false))
            .withTemplate(IssueSxssfTransformerTest.class.getResourceAsStream(getClass().getSimpleName() + ".xlsx"))
        );

        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet(1);
            assertEquals("Manager:", w.getCellValueAsString(3, 1)); // A3
            assertEquals("ABC", w.getCellValueAsString(3, 2)); // B3
        }
    }

    private Context prepareContext() {
        final Context context = new Context();
        context.setFormulaProcessingRequired(false);

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
    
    public class MyJxlsTemplateFiller extends JxlsTemplateFiller {
        
        protected MyJxlsTemplateFiller(JxlsOptions options, InputStream template) {
            super(options, template);
        }

        @Override
        protected void processAreas(Map<String, Object> data) {
            areas = options.getAreaBuilder().build(transformer, true);
            for (Area area : areas) {
                CellRef ref = new CellRef("Result", 0, 0);
                area.applyAt(ref, new Context(null, data));
            }
        }
    }
}
