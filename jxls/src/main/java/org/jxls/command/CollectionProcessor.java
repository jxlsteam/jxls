package org.jxls.command;

import java.util.List;

import org.jxls.area.Area;
import org.jxls.command.EachCommand.Direction;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.JxlsException;
import org.jxls.common.Size;
import org.jxls.expression.ExpressionEvaluator;
import org.jxls.transform.TransformationConfig;
import org.jxls.util.JxlsHelper;
import org.jxls.util.UtilWrapper;

/**
 * Traverse collection for EachCommand
 */
public class CollectionProcessor {
    private final Context context;
    private final UtilWrapper util;
    private final String varIndex;
    private final EachCommand.Direction direction;
    private final String select;
    private final String groupBy;
    private final String multisheet;
    private CellRefGenerator cellRefGenerator;
    private final Area area;

    private CellRef currentCell;
    private int index;
    private int newWidth;
    private int newHeight;
    
    public CollectionProcessor(Context context, UtilWrapper util, String varIndex, Direction direction, String select,
            String groupBy, String multisheet, CellRefGenerator cellRefGenerator, Area area) {
        this.context = context;
        this.util = util;
        this.varIndex = varIndex;
        this.direction = direction;
        this.select = select;
        this.groupBy = groupBy;
        this.multisheet = multisheet;
        this.cellRefGenerator = cellRefGenerator;
        this.area = area;
    }
    
    public interface TransformationConfigGetter {
        TransformationConfig get();
    }

    public void initMultiSheet(CellRef cellRef, TransformationConfigGetter transformationConfigGetter) {
        if (cellRefGenerator == null && multisheet != null) {
            List<String> sheetNameList = extractSheetNameList();
            cellRefGenerator = sheetNameList == null
                    ? new DynamicSheetNameGenerator(multisheet, cellRef, transformationConfigGetter.get().getExpressionEvaluator())
                    : new SheetNameGenerator(sheetNameList, cellRef);
        }
    }

    @SuppressWarnings("unchecked")
    private List<String> extractSheetNameList() {
        try {
            Object sheetnames = context.getVar(multisheet);
            if (sheetnames == null) {
                return null;
            } else if (sheetnames instanceof List) {
                return (List<String>) sheetnames;
            }
        } catch (Exception e) {
            throw new JxlsException("Failed to get sheet names from " + multisheet, e);
        }
        throw new JxlsException("The sheet names var '" + multisheet + "' must be of type List<String>.");
    }

    public Size processCollection(Iterable<?> itemsCollection, CellRef cellRef, String varName) {
        index = 0;
        newWidth = 0;
        newHeight = 0;

        ExpressionEvaluator selectEvaluator = null;
        if (select != null && (groupBy == null || groupBy.isEmpty())) {
            selectEvaluator = JxlsHelper.getInstance().createExpressionEvaluator(select);
        }

        currentCell = cellRef;
        Object currentVarObject = varName == null ? null : context.getRunVar(varName);
        Object currentVarIndexObject = varIndex == null ? null : context.getRunVar(varIndex);
        int currentIndex = 0;
        for (Object obj : itemsCollection) {
            context.putVar(varName, obj);
            if (varIndex != null) {
                context.putVar(varIndex, currentIndex);
            }
            if (selectEvaluator == null || util.isConditionTrue(selectEvaluator, context)) {
                if (processItem(obj)) {
                    break;
                }
                currentIndex++;
            }
        }
        restoreVarObject(varIndex, currentVarIndexObject);
        restoreVarObject(varName, currentVarObject);
        
        return new Size(newWidth, newHeight);
    }

    protected boolean processItem(Object item) {
        if (cellRefGenerator != null) {
            currentCell = cellRefGenerator.generateCellRef(index++, context);
        }
        if (currentCell == null) {
            return true; // break loop
        }
        Size size;
        try {
            size = area.applyAt(currentCell, context);
        } catch (NegativeArraySizeException e) {
            throw new JxlsException("Check jx:each/lastCell parameter in template! Illegal area: " + area.getAreaRef(), e);
        }
        if (cellRefGenerator != null) {
            newWidth = Math.max(newWidth, size.getWidth());
            newHeight = Math.max(newHeight, size.getHeight());
        } else if (direction == Direction.DOWN) {
            currentCell = new CellRef(currentCell.getSheetName(), currentCell.getRow() + size.getHeight(), currentCell.getCol());
            newWidth = Math.max(newWidth, size.getWidth());
            newHeight += size.getHeight();
        } else { // RIGHT
            currentCell = new CellRef(currentCell.getSheetName(), currentCell.getRow(), currentCell.getCol() + size.getWidth());
            newWidth += size.getWidth();
            newHeight = Math.max(newHeight, size.getHeight());
        }
        return false;
    }

    private void restoreVarObject(String varName, Object varObject) {
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
