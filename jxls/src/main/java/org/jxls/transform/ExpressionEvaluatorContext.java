package org.jxls.transform;

import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.jxls.builder.JxlsTemplateFillerBuilder;
import org.jxls.common.EvaluationResult;
import org.jxls.expression.ExpressionEvaluator;
import org.jxls.expression.ExpressionEvaluatorFactory;

public class ExpressionEvaluatorContext {
    private static final String EXPRESSION_PART = "(.+?)";

    private final ExpressionEvaluatorFactory expressionEvaluatorFactory;
    private final String expressionNotationBegin;
    private final String expressionNotationEnd;
    private final Pattern expressionNotationPattern;
    private ExpressionEvaluator expressionEvaluator = null;

    public ExpressionEvaluatorContext(ExpressionEvaluatorFactory expressionEvaluatorFactory, String expressionNotationBegin, String expressionNotationEnd) {
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

    /**
     * INTERNAL
     * @param rawExpression e.g. "${e.name}"
     * @param data -
     * @return EvaluationResult
     */
    public EvaluationResult evaluateRawExpression(String rawExpression, Map<String, Object> data) {
        StringBuilder sb = new StringBuilder();
        int beginExpressionLength = expressionNotationBegin.length();
        int endExpressionLength = expressionNotationEnd.length();
        Matcher exprMatcher = expressionNotationPattern.matcher(rawExpression);
        ExpressionEvaluator evaluator = getExpressionEvaluator();
        String matchedString;
        String expression;
        Object lastMatchEvalResult = null;
        int matchCount = 0;
        int endOffset = 0;
        while (exprMatcher.find()) {
            endOffset = exprMatcher.end();
            matchCount++;
            matchedString = exprMatcher.group();
            expression = matchedString.substring(beginExpressionLength, matchedString.length() - endExpressionLength);
            lastMatchEvalResult = evaluator.evaluate(expression, data);
            exprMatcher.appendReplacement(sb,
                    Matcher.quoteReplacement(lastMatchEvalResult != null ? lastMatchEvalResult.toString() : ""));
        }
        String lastStringResult = lastMatchEvalResult != null ? lastMatchEvalResult.toString() : "";
        boolean isAppendTail = matchCount == 1 && endOffset < rawExpression.length();
        if (matchCount > 1 || isAppendTail) {
            exprMatcher.appendTail(sb);
            return new EvaluationResult(sb.toString());
        } else if (matchCount == 1) {
            if (sb.length() > lastStringResult.length()) {
                return new EvaluationResult(sb.toString());
            } else {
                return new EvaluationResult(lastMatchEvalResult, true);
            }
        } else if (matchCount == 0) {
            return new EvaluationResult(rawExpression);
        }
        return null;
    }
}
