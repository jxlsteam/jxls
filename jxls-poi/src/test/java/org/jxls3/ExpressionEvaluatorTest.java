package org.jxls3;

import static org.junit.Assert.assertTrue;

import java.util.HashMap;

import org.junit.Assert;
import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.common.Context;
import org.jxls.common.JxlsException;
import org.jxls.common.PoiExceptionLogger;
import org.jxls.common.PoiExceptionThrower;
import org.jxls.expression.EvaluationException;
import org.jxls.expression.ExpressionEvaluatorFactoryJSR223Impl;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

public class ExpressionEvaluatorTest {
    private boolean catched;

    @Test
    public void wrongExpression() {
        try {
            Jxls3Tester.xlsx(getClass()).test(new HashMap<>(),
                    JxlsPoiTemplateFillerBuilder.newInstance().withLogger(new PoiExceptionThrower()));
            Assert.fail("JxlsException expected");
        } catch (JxlsException e) {
            assertTrue(e.getCause().getCause() instanceof EvaluationException);
        }
    }

    @Test
    public void wrongExpressionSwallowed() {
        catched = false;
        Jxls3Tester.xlsx(getClass()).test(new HashMap<>(),
                JxlsPoiTemplateFillerBuilder.newInstance().withLogger(new PoiExceptionLogger() {
                    @Override
                    public void handleCellException(Exception e, String cell, Context context) {
                        catched = "CellData{Sheet1!A2, cellType=STRING, cellValue=${a***3}}".equals(cell);
                    }
                }));
        assertTrue(catched);
    }

    @Test
    public void noJSR223ScriptEngine() {
        try {
            Jxls3Tester.xlsx(getClass()).test(new HashMap<>(),
                    JxlsPoiTemplateFillerBuilder.newInstance()
                    .withExpressionEvaluatorFactory(new ExpressionEvaluatorFactoryJSR223Impl("nonsense"))
                    .withLogger(new PoiExceptionThrower()));
            Assert.fail("JxlsException expected");
        } catch (JxlsException e) {
            assertTrue(e.getMessage().contains("Can not get script engine"));
        }
    }

    @Test
    public void wrongExpression_JSR223() {
        try {
            Jxls3Tester.xlsx(getClass()).test(new HashMap<>(),
                    JxlsPoiTemplateFillerBuilder.newInstance()
                    .withExpressionEvaluatorFactory(new ExpressionEvaluatorFactoryJSR223Impl("JEXL3"))
                    .withLogger(new PoiExceptionThrower()));
            Assert.fail("JxlsException expected");
        } catch (JxlsException e) {
            assertTrue(e.getCause().getCause() instanceof EvaluationException);
        }
    }
}
