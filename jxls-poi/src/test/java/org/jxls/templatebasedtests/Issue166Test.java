package org.jxls.templatebasedtests;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.JxlsTester.TransformerChecker;
import org.jxls.area.Area;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.formula.StandardFormulaProcessor;
import org.jxls.transform.Transformer;

/**
 * Wrong average on 2nd sheet
 */
public class Issue166Test {

    @Test
    public void test() {
    	// Prepare: define result set
        List<Map<String, Object>> rs = new ArrayList<>();
        for (int i = 0; i < 5; i++) {
            Map<String, Object> map = new HashMap<String, Object>();
            map.put("count", i);
            rs.add(map);
        }
        final Context context = new Context();
        context.putVar("rs0", rs);
        
        TransformerChecker myProcessing = new TransformerChecker() {
            @Override
            public Transformer checkTransformer(Transformer transformer) {
                transformer.setEvaluateFormulas(false);
                AreaBuilder areaBuilder = new XlsCommentAreaBuilder();
                areaBuilder.setTransformer(transformer);
                List<Area> xlsAreaList = areaBuilder.build();
                for (Area xlsArea : xlsAreaList) {
                    xlsArea.applyAt(new CellRef(xlsArea.getStartCellRef().getCellName()), context);
                    // Make sure the StandardFormulaProcessor is used. This will make sure formula
                    // references on multi-sheet workbooks are correct.
                    xlsArea.setFormulaProcessor(new StandardFormulaProcessor());
                    xlsArea.processFormulas();
                }
                return transformer;
            }
        };

        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.createTransformerAndProcessTemplate(context, myProcessing);
        
        // Verify
        // TODO assertions
    }
}
