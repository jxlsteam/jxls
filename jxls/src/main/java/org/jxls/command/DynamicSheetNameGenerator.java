package org.jxls.command;

import java.util.HashSet;
import java.util.Set;

import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.expression.ExpressionEvaluator;
import org.jxls.transform.SafeSheetNameBuilder;

/**
 * Creates cell references based on passed sheet names. Appends unique number to the name if name already exists.
 */
public class DynamicSheetNameGenerator implements CellRefGenerator {
    private final String sheetNameExpression;
    private final CellRef startCellRef;
    private final ExpressionEvaluator expressionEvaluator;
    private final Set<String> usedSheetNames = new HashSet<String>(); // only used if there's no SafeSheetNameBuilder

    public DynamicSheetNameGenerator(String sheetNameExpression, CellRef startCellRef, ExpressionEvaluator expressionEvaluator) {
        this.sheetNameExpression = sheetNameExpression;
        this.startCellRef = startCellRef;
        this.expressionEvaluator = expressionEvaluator;
    }

    @Override
    public CellRef generateCellRef(int index, Context context) {
        String sheetName = (String) expressionEvaluator.evaluate(sheetNameExpression, context.toMap());
        boolean safeName = false;
        Object builder = context.getVar(SafeSheetNameBuilder.CONTEXT_VAR_NAME);
        if (builder instanceof SafeSheetNameBuilder) {
            // The SafeSheetNameBuilder builds a valid and unique sheetName. This is the new style.
            sheetName = ((SafeSheetNameBuilder) builder).createSafeSheetName(sheetName, index);
            safeName = true;
        }
        if (sheetName == null) {
            return null;
        }
        if (!safeName && !usedSheetNames.add(sheetName)) {
            // This is the old-style algorithm for backward compatibility.
            // name already used
            for (int i = 1;; i++) {
                String newName = sheetName + '(' + i + ')';
                if (usedSheetNames.add(newName)) {
                    sheetName = newName;
                    break;
                } // else: the name is already used, continue
            }
        }
        return new CellRef(sheetName, startCellRef.getRow(), startCellRef.getCol());
    }
}
