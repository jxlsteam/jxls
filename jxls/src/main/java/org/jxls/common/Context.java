package org.jxls.common;

import java.util.Map;

import org.jxls.expression.ExpressionEvaluator;

/**
 * Jxls context interface (Jxls internal class)
 */
public interface Context extends PublicContext {

    Map<String, Object> toMap();

    ExpressionEvaluator getExpressionEvaluator(String expression);
    
    EvaluationResult _evaluateRawExpression(String rawExpression);

    boolean isFormulaProcessingRequired();

    void setFormulaProcessingRequired(boolean formulaProcessingRequired);

    boolean isIgnoreSourceCellStyle();

    void setIgnoreSourceCellStyle(boolean ignoreSourceCellStyle);

    Map<String, String> getCellStyleMap();

    void setCellStyleMap(Map<String, String> cellStyleMap);
}
