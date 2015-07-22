package org.jxls.transform;

import org.jxls.expression.ExpressionEvaluator;
import org.jxls.expression.JexlExpressionEvaluator;

import java.util.regex.Pattern;

/**
 * Created by Leonid Vysochyn on 22-Jul-15.
 */
public class TransformationConfig {
    public static final String EXPRESSION_PART = "(.+?)";
    public static final String DEFAULT_EXPRESSION_BEGIN = "${";
    public static final String DEFAULT_EXPRESSION_END = "}";

    ExpressionEvaluator expressionEvaluator = new JexlExpressionEvaluator();
    String expressionNotationBegin = DEFAULT_EXPRESSION_BEGIN;
    String expressionNotationEnd = DEFAULT_EXPRESSION_END;
    Pattern expressionNotationPattern;

    {
        buildExpressionNotation(DEFAULT_EXPRESSION_BEGIN, DEFAULT_EXPRESSION_END);
    }

    public TransformationConfig() {
    }

    void buildExpressionNotation(String expressionBegin, String expressionEnd){
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

    public Pattern getExpressionNotationPattern(){
        return expressionNotationPattern;
    }
}
