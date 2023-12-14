package org.jxls.transform;

import java.util.regex.Pattern;

import org.jxls.builder.JxlsTemplateFillerBuilder;
import org.jxls.expression.ExpressionEvaluator;
import org.jxls.expression.ExpressionEvaluatorFactory;
import org.jxls.util.JxlsHelper;

/**
 * Transformation configuration class
 */
public class TransformationConfig {
    private static final String EXPRESSION_PART = "(.+?)";
    private static final String DEFAULT_EXPRESSION_BEGIN = JxlsTemplateFillerBuilder.DEFAULT_EXPRESSION_BEGIN;
    private static final String DEFAULT_EXPRESSION_END = JxlsTemplateFillerBuilder.DEFAULT_EXPRESSION_END;
    private static final String DEFAULT_REGEX_EXPRESSION = "\\$\\{[^}]*}";

    private ExpressionEvaluatorFactory expressionEvaluatorFactory;
    private ExpressionEvaluator expressionEvaluator;
    private String expressionNotationBegin = DEFAULT_EXPRESSION_BEGIN;
    private String expressionNotationEnd = DEFAULT_EXPRESSION_END;
    private Pattern expressionNotationPattern = Pattern.compile(DEFAULT_REGEX_EXPRESSION);

    public void buildExpressionNotation(String expressionBegin, String expressionEnd) {
        this.expressionNotationBegin = expressionBegin;
        this.expressionNotationEnd = expressionEnd;
        String regexExpression = Pattern.quote(expressionNotationBegin) + EXPRESSION_PART + Pattern.quote(expressionNotationEnd);
        expressionNotationPattern = Pattern.compile(regexExpression);
    }

    public void setExpressionEvaluatorFactory(ExpressionEvaluatorFactory expressionEvaluatorFactory) {
        this.expressionEvaluatorFactory = expressionEvaluatorFactory;
        expressionEvaluator = null;
    }

    public ExpressionEvaluator getExpressionEvaluator() {
		if (expressionEvaluator == null) {
			if (expressionEvaluatorFactory == null) {
				expressionEvaluatorFactory = JxlsHelper.getInstance().getExpressionEvaluatorFactory();
			}
			expressionEvaluator = expressionEvaluatorFactory.createExpressionEvaluator(null);
		}
		return expressionEvaluator;
    }

    public ExpressionEvaluator getExpressionEvaluator(String expression) {
		if (expressionEvaluatorFactory == null) {
			expressionEvaluatorFactory = JxlsHelper.getInstance().getExpressionEvaluatorFactory();
		}
		return expressionEvaluatorFactory.createExpressionEvaluator(expression);
    }

    public String getExpressionNotationBegin() {
        return expressionNotationBegin;
    }

    public String getExpressionNotationEnd() {
        return expressionNotationEnd;
    }

    public Pattern getExpressionNotationPattern() {
        return expressionNotationPattern;
    }
}
