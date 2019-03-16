package org.jxls.command;

import java.util.HashSet;
import java.util.Set;

import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.expression.ExpressionEvaluator;

/**
 * Creates cell references based on passed sheet names. Appends unique number to the name if name already exists.
 */
public class DynamicSheetNameGenerator implements CellRefGenerator {
    private final Set<String> names = new HashSet<String>();
    private final String sheetName;
    private final CellRef startCellRef;
    private final ExpressionEvaluator expressionEvaluator;

    public DynamicSheetNameGenerator(String sheetName, CellRef startCellRef, ExpressionEvaluator expressionEvaluator) {
        this.sheetName = sheetName;
        this.startCellRef = startCellRef;
        this.expressionEvaluator = expressionEvaluator;
    }

    @Override
    public CellRef generateCellRef(int index, Context context) {
        String name = (String) expressionEvaluator.evaluate(sheetName, context.toMap());
        if (name == null) {
            return null;
        }
        if (!names.add(name)) {
            // name already used
            for (int i = 1;; i++) {
                String tmp = name + '(' + i + ')';
                if (names.add(tmp)) {
                    name = tmp;
                    break;
                } // else: the name is already used, continue
            }
        }
        return new CellRef(name, startCellRef.getRow(), startCellRef.getCol());
    }
}
