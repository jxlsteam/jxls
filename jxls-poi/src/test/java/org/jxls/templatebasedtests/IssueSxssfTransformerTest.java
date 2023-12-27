package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Map;

import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.TestWorkbook;
import org.jxls.area.Area;
import org.jxls.area.XlsArea;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.JxlsStreaming;
import org.jxls.builder.JxlsTemplateFiller;
import org.jxls.builder.KeepTemplateSheet;
import org.jxls.command.Command;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.PoiExceptionThrower;
import org.jxls.expression.ExpressionEvaluatorFactory;
import org.jxls.formula.FormulaProcessor;
import org.jxls.logging.JxlsLogger;
import org.jxls.transform.JxlsTransformerFactory;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;
import org.jxls.transform.poi.PoiContext;
import org.jxls.transform.poi.PoiTransformer;
import org.jxls.transform.poi.PoiTransformerFactory;

/**
 * Test for issue 153
 * 
 * Issue in Excel Output while using SXSSF Transformer with JXLS >= 2.7.0
 * cause: commit 5354beaf
 */
public class IssueSxssfTransformerTest {
    
    @Test
    public void test() throws IOException {
        PoiTransformerFactory transformerFactory = new PoiTransformerFactory() {
            @Override
            protected PoiTransformer createTransformer(Workbook workbook, JxlsStreaming streaming) {
                return PoiTransformer.createSxssfTransformer(workbook, 1000, true);
            };
        };

        Jxls3Tester tester = Jxls3Tester.xlsx(getClass());
        tester.test(prepareContext().toMap(), new JxlsPoiTemplateFillerBuilder() {
            @Override
            public JxlsTemplateFiller build() {
                return new MyJxlsTemplateFiller(getExpressionEvaluatorFactory(), expressionNotationBegin,
                        expressionNotationEnd, new PoiExceptionThrower(), getFormulaProcessor(), ignoreColumnProps,
                        ignoreRowProps, recalculateFormulasBeforeSaving, recalculateFormulasOnOpening,
                        keepTemplateSheet, getAreaBuilder(), commands, clearTemplateCells, transformerFactory, JxlsStreaming.STREAMING_ON,
                        IssueSxssfTransformerTest.class.getResourceAsStream("IssueSxssfTransformerTest.xlsx"));
            }
        }.withRecalculateFormulasOnOpening(true));

        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet(1);
            assertEquals("Manager:", w.getCellValueAsString(3, 1)); // A3
            assertEquals("ABC", w.getCellValueAsString(3, 2)); // B3
        }
    }

    private Context prepareContext() {
        final Context context = new PoiContext();
        context.putVar(XlsArea.IS_FORMULA_PROCESSING_REQUIRED, Boolean.FALSE);

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

        protected MyJxlsTemplateFiller(ExpressionEvaluatorFactory expressionEvaluatorFactory,
                String expressionNotationBegin, String expressionNotationEnd, JxlsLogger logger,
                FormulaProcessor formulaProcessor, boolean ignoreColumProps, boolean ignoreRowProps,
                boolean recalculateFormulasBeforeSaving, boolean recalculateFormulasOnOpening,
                KeepTemplateSheet keepTemplateSheet, AreaBuilder areaBuilder,
                Map<String, Class<? extends Command>> commands, boolean clearTemplateCells,
                JxlsTransformerFactory transformerFactory, JxlsStreaming streaming, InputStream template) {
            super(expressionEvaluatorFactory, expressionNotationBegin, expressionNotationEnd, logger, formulaProcessor,
                    ignoreColumProps, ignoreRowProps, recalculateFormulasBeforeSaving, recalculateFormulasOnOpening,
                    keepTemplateSheet, areaBuilder, commands, clearTemplateCells, transformerFactory, streaming, template);
        }
        
        @Override
        protected void processAreas(Map<String, Object> data) {
            areas = areaBuilder.build(transformer, true);
            for (Area area : areas) {
                CellRef ref = new CellRef("Result", 0, 0);
                area.applyAt(ref, new Context(data));
            }
//            workbook.setActiveSheet(activeSheetIndex);
//            SXSSFWorkbook workbook2 = (SXSSFWorkbook) transformer.getWorkbook();
//            workbook2.write(os);
        }
    }
}
