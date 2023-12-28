package org.jxls.transform;

import java.util.regex.Pattern;

import org.jxls.builder.JxlsTemplateFillerBuilder;
import org.jxls.expression.ExpressionEvaluator;
import org.jxls.expression.ExpressionEvaluatorFactory;

public class TransformationConfig { // TODO find better name
    private static final String EXPRESSION_PART = "(.+?)";

    private final ExpressionEvaluatorFactory expressionEvaluatorFactory;
    private final String expressionNotationBegin;
    private final String expressionNotationEnd;
    private final Pattern expressionNotationPattern;
    private ExpressionEvaluator expressionEvaluator = null;

    public TransformationConfig(ExpressionEvaluatorFactory expressionEvaluatorFactory, String expressionNotationBegin, String expressionNotationEnd) {
        if (expressionEvaluatorFactory == null) {
            throw new IllegalArgumentException("expressionEvaluatorFactory must not be null");
        }
        this.expressionEvaluatorFactory = expressionEvaluatorFactory;
        this.expressionNotationBegin = expressionNotationBegin == null ? JxlsTemplateFillerBuilder.DEFAULT_EXPRESSION_BEGIN : expressionNotationBegin;
        this.expressionNotationEnd = expressionNotationEnd == null ? JxlsTemplateFillerBuilder.DEFAULT_EXPRESSION_END : expressionNotationEnd;

        expressionNotationPattern = Pattern.compile(Pattern.quote(this.expressionNotationBegin) + EXPRESSION_PART + Pattern.quote(this.expressionNotationEnd));
    }
    
    public ExpressionEvaluator getExpressionEvaluator() {
		if (expressionEvaluator == null) {
			expressionEvaluator = expressionEvaluatorFactory.createExpressionEvaluator(null);
		}
		return expressionEvaluator;
    }

    public ExpressionEvaluator getExpressionEvaluator(String expression) {
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
