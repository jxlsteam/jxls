package org.jxls.transform;

import java.util.regex.Pattern;

import org.jxls.expression.ExpressionEvaluator;
import org.jxls.util.JxlsHelper;

/**
 * Transformation configuration class
 */
public class TransformationConfig {
    private static final String EXPRESSION_PART = "(.+?)";
    private static final String DEFAULT_EXPRESSION_BEGIN = "${";
    private static final String DEFAULT_EXPRESSION_END = "}";
    private static final String DEFAULT_REGEX_EXPRESSION = "\\$\\{[^}]*}";

    private ExpressionEvaluator expressionEvaluator =  JxlsHelper.getInstance().createExpressionEvaluator(null);
    private String expressionNotationBegin = DEFAULT_EXPRESSION_BEGIN;
    private String expressionNotationEnd = DEFAULT_EXPRESSION_END;
    private Pattern expressionNotationPattern = Pattern.compile(DEFAULT_REGEX_EXPRESSION);

    public void buildExpressionNotation(String expressionBegin, String expressionEnd) {
        this.expressionNotationBegin = expressionBegin;
        this.expressionNotationEnd = expressionEnd;
        String regexExpression = Pattern.quote(expressionNotationBegin) + EXPRESSION_PART + Pattern.quote(expressionNotationEnd);
        expressionNotationPattern = Pattern.compile(regexExpression);
    }

    public ExpressionEvaluator getExpressionEvaluator() {
        return expressionEvaluator;
    }

    public void setExpressionEvaluator(ExpressionEvaluator expressionEvaluator) {
        this.expressionEvaluator = expressionEvaluator;
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
