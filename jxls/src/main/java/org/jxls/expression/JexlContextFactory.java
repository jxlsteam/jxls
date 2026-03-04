package org.jxls.expression;

import java.util.Map;

import org.apache.commons.jexl3.JexlContext;

public interface JexlContextFactory {

	JexlContext create(Map<String, Object> context);
}
