package org.jxls.expression;

import javax.script.ScriptEngine;
import javax.script.ScriptEngineManager;

/**
 * A factory to create {@link ExpressionEvaluator} instance which is based on JSR 223
 */
public final class ExpressionEvaluatorFactoryJSR223Impl implements ExpressionEvaluatorFactory {
    private final ScriptEngineManager manager = new ScriptEngineManager();
    private final ScriptEngine scriptEngine;

    public ExpressionEvaluatorFactoryJSR223Impl(String lang) {
        scriptEngine = manager.getEngineByName(lang);
    }

    @Override
    public ExpressionEvaluator createExpressionEvaluator(final String expression) {
        return new ExpressionEvaluator4JSR223Impl(scriptEngine, expression);
    }
}
