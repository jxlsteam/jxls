package org.jxls.templatebasedtests;

import static org.junit.Assert.assertTrue;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.JxlsTester.TransformerChecker;
import org.jxls.common.Context;
import org.jxls.transform.Transformer;
import org.jxls.transform.poi.PoiTransformer;
import org.jxls.transformer.OwnTransformer;

/**
 * This testcase ensures that PoiTransformer is replaceable or extendable.
 */
public class UseOwnTransformerTest {
    private OwnTransformer own;
    
    @Test
    public void testReplacementOfPoiTransformer() {
        TransformerChecker useOwnTransformer = new TransformerChecker() {
            @Override
            public Transformer checkTransformer(Transformer transformer) {
                own = new OwnTransformer(((PoiTransformer) transformer).getWorkbook());
                own.setOutputStream(((PoiTransformer) transformer).getOutputStream());
                return own;
            }
        };
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.createTransformerAndProcessTemplate(new Context(), useOwnTransformer);
        
        assertTrue("clearCell() must be called", own.isClearCellCalled());
        assertTrue("transform() must be called", own.isTransformCalled());
    }
}
