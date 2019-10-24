package org.jxls.expression;

import java.util.Map;

import javax.script.Bindings;
import javax.script.ScriptEngine;
import javax.script.ScriptException;
import javax.script.SimpleBindings;

/**
 * JSR 223 based implementation of the {@link ExpressionEvaluator} interface
 */
public class ExpressionEvaluator4JSR223Impl implements ExpressionEvaluator {
	private final String expression;
	private final ScriptEngine scriptEngine;
	
	public ExpressionEvaluator4JSR223Impl(ScriptEngine scriptEngine, String expression) {
		this.scriptEngine = scriptEngine;
		this.expression = expression;
	}
	
    private static class BindingCacheHolder {
        Bindings binding;
        Map<String, Object> context;
    }

    private static final ThreadLocal<BindingCacheHolder> threadLocalCache = new ThreadLocal<BindingCacheHolder>() {
        
        @Override
        protected BindingCacheHolder initialValue() {
            return new BindingCacheHolder();
        }
    };

    @Override
    public Object evaluate(final String expression, final Map<String, Object> context) {
        if (expression == null || context == null) {
            return null;
        }
        final BindingCacheHolder holder = threadLocalCache.get();
        if (holder.binding == null) {
            holder.binding = new SimpleBindings(context);
        }
        if (holder.context == null || holder.context != context) {
            holder.context = context;
            holder.binding.putAll(context);
        }
        try {
            return scriptEngine.eval(expression, holder.binding);
        } catch (ScriptException e) {
            throw new EvaluationException("Evaluate error on: " + expression, e);
        }
    }

	@Override
	public Object evaluate(Map<String, Object> context) {
		return evaluate(expression,context);
	}

	@Override
	public String getExpression() {
		return expression;
	}
}
