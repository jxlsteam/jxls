package org.jxls.transform;

import java.util.Map;

import org.jxls.common.EvaluationResult;
import org.jxls.expression.ExpressionEvaluator;
import org.jxls.expression.ExpressionEvaluatorFactory;

/**
 * ExpressionEvaluatorContext implementation that allows ${..} expressions inside ${..} if they are contained within double quotes.
 */
public class ExpressionEvaluatorContextQ extends ExpressionEvaluatorContext {

    public ExpressionEvaluatorContextQ(ExpressionEvaluatorFactory expressionEvaluatorFactory, String expressionNotationBegin, String expressionNotationEnd) {
        super(expressionEvaluatorFactory, expressionNotationBegin, expressionNotationEnd);
    }

    @Override
    public EvaluationResult evaluateRawExpression(String rawExpression, Map<String, Object> data) {
        final boolean one = rawExpression.startsWith(expressionNotationBegin) && rawExpression.endsWith(expressionNotationEnd);
        final ExpressionEvaluator evaluator = getExpressionEvaluator();
        int replacements = 0;
        Object lastMatchEvalResult = null;
        int o = rawExpression.indexOf(expressionNotationBegin);
        while (o >= 0) {
            int nextStart = o + expressionNotationBegin.length();
            int oo = findEnd(rawExpression, nextStart);
            if (oo >= 0) {
                String expr = rawExpression.substring(nextStart, oo);
                lastMatchEvalResult = evaluator.evaluate(expr, data);
                String newRawExpression = rawExpression.substring(0, o);
                if (lastMatchEvalResult != null) {
                    newRawExpression += lastMatchEvalResult;
                }
                nextStart = newRawExpression.length();
                rawExpression = newRawExpression + rawExpression.substring(oo + expressionNotationEnd.length());
                replacements++;
            }

            o = rawExpression.indexOf(expressionNotationBegin, nextStart);
        }
        if (one && replacements == 1) {
            return new EvaluationResult(lastMatchEvalResult, true);
        }
        return new EvaluationResult(rawExpression);
    }
    
    private int findEnd(String rawExpression, final int start0) {
        int start = start0;
        int end = -1;
        boolean again;
        do {
            again = false;
            end = rawExpression.indexOf(expressionNotationEnd, start);
            if (end > -1) {
                boolean inQuotes = false;
                for (int i = start0; i < end; i++) {
                    char c = rawExpression.charAt(i);
                    if (c == '"' && !(i - 1 >= 0 && rawExpression.charAt(i - 1) == '\\')) {
                        inQuotes = !inQuotes;
                    }
                }
                if (inQuotes) {
                    start++; // expressionNotationEnd inside quotes will be ignored
                    again = true;
                }
            }
        } while (again);
        return end;
    }
}
