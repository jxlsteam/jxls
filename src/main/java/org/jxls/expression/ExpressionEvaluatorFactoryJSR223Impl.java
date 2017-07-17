package org.jxls.expression;

import javax.script.ScriptEngine;
import javax.script.ScriptEngineManager;

import org.jxls.util.JxlsHelper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public final class ExpressionEvaluatorFactoryJSR223Impl implements ExpressionEvaluatorFactory {
	
	/**
	 * Logger for this class
	 */
	private static final Logger LOG = LoggerFactory.getLogger(ExpressionEvaluatorFactoryJSR223Impl.class);
	
	private final ScriptEngineManager manager = new ScriptEngineManager();
	private ScriptEngine scriptEngine;
	public ExpressionEvaluatorFactoryJSR223Impl() {
		super();
		final String lang = JxlsHelper.getProperty("jxls.script_engine", "jexl");
		LOG.info("jxls.script_engine:{}",lang);
		scriptEngine = manager.getEngineByName(lang.trim());
	}

	@Override
	public ExpressionEvaluator createExpressionEvaluator(final String expression) {
		return new ExpressionEvaluator4JSR223Impl(scriptEngine,expression);
	}

}
