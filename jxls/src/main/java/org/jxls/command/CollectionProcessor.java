package org.jxls.command;

import org.jxls.common.Context;
import org.jxls.expression.ExpressionEvaluator;

public abstract class CollectionProcessor {

    public void traverse(Context context, Iterable<?> itemsCollection, String varName, String varIndex,
            ExpressionEvaluator selectEvaluator) {
        Object currentVarObject = EachCommand.getRunVar(context, varName);
        Object currentVarIndexObject = EachCommand.getRunVar(context, varIndex);
        int currentIndex = 0;
        for (Object obj : itemsCollection) {
            context.putVar(varName, obj);
            if (varIndex != null) {
                context.putVar(varIndex, Integer.valueOf(currentIndex));
            }
            if (selectEvaluator != null && !Boolean.TRUE.equals(selectEvaluator.isConditionTrue(context.toMap()))) {
                continue;
            }
            if (processItem(obj)) {
                break;
            }
            currentIndex++;
        }
        restoreVarObject(context, varIndex, currentVarIndexObject);
        restoreVarObject(context, varName, currentVarObject);
    }
    
    /**
     * @param obj -
     * @return true: break, false: go on
     */
    protected abstract boolean processItem(Object obj);

    private void restoreVarObject(Context context, String varName, Object varObject) {
        if (varName == null) {
            return;
        }
        if (varObject != null) {
            context.putVar(varName, varObject);
        } else {
            context.removeVar(varName);
        }
    }
}
