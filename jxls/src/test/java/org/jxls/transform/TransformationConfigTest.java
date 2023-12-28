package org.jxls.transform;

import static org.junit.Assert.assertEquals;

import org.junit.Test;
import org.jxls.common.Context;
import org.jxls.expression.ExpressionEvaluatorFactoryJexlImpl;

/**
 * @author Leonid Vysochyn
 */
public class TransformationConfigTest {
    
    @Test
    public void defaultValuesAfterCreation() {
        TransformationConfig config = new Context().getTransformationConfig();
        assertEquals("Default expression notation begin part is wrong", "${", config.getExpressionNotationBegin());
        assertEquals("Default expression notation end part is wrong", "}", config.getExpressionNotationEnd());
    }

    @Test
    public void buildExpressionNotation() {
        String expressionBegin = "[[";
        String expressionEnd = "]]";
        TransformationConfig config = new TransformationConfig(new ExpressionEvaluatorFactoryJexlImpl(), expressionBegin, expressionEnd);
        assertEquals("Expression notation begin part is wrong after buildExpressionNotation()", expressionBegin, config.getExpressionNotationBegin());
        assertEquals("Expression notation end part is wrong after buildExpressionNotation()", expressionEnd, config.getExpressionNotationEnd());
    }
}
