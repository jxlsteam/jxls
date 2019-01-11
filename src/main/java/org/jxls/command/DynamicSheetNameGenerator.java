package org.jxls.command;

import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.expression.ExpressionEvaluator;

/**
 * Creates cell references based on passed sheet names
 */
public class DynamicSheetNameGenerator implements CellRefGenerator {
    private String sheetName;
    private CellRef startCellRef;
    private ExpressionEvaluator expressionEvaluator;

    public DynamicSheetNameGenerator(String sheetName, CellRef startCellRef, ExpressionEvaluator expressionEvaluator) {
        this.sheetName = sheetName;
        this.startCellRef = startCellRef;
        this.expressionEvaluator = expressionEvaluator;
    }

    @Override
    public CellRef generateCellRef(int index, Context context) {
        String name = (String)expressionEvaluator.evaluate(sheetName, context.toMap());
        if (name == null) {
            return null;
        }
        return new CellRef(name, startCellRef.getRow(), startCellRef.getCol());
    }
}
