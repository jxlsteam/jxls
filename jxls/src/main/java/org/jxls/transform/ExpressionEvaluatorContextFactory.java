package org.jxls.transform;

import org.jxls.expression.ExpressionEvaluatorFactory;

public interface ExpressionEvaluatorContextFactory {

    ExpressionEvaluatorContext build(ExpressionEvaluatorFactory expressionEvaluatorFactory, String expressionNotationBegin, String expressionNotationEnd);
}
