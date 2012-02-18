package com.jxls.writer.command;

import com.jxls.writer.area.Area;
import com.jxls.writer.common.CellRef;
import com.jxls.writer.common.Context;
import com.jxls.writer.common.Size;
import com.jxls.writer.expression.ExpressionEvaluator;
import com.jxls.writer.expression.JexlExpressionEvaluator;

import java.util.Collection;

/**
 * @author Leonid Vysochyn
 *         Date: 2/17/12 3:02 PM
 */
public class EachCellCommand extends AbstractCommand {

    String var;
    String items;
    CellRefGenerator cellRefGenerator;
    Area area;

    public EachCellCommand(String var, String items, CellRefGenerator cellRefGenerator, Area area) {
        this.var = var;
        this.items = items;
        this.cellRefGenerator = cellRefGenerator;
        this.area = area;
    }

    public String getName() {
        return "EachSheet";
    }

    public Size applyAt(CellRef cellRef, Context context) {
        Collection itemsCollection = calculateItemsCollection(context);
        CellRef currentCell;
        int width = 0;
        int height = 0;
        int index = 0;
        for( Object obj : itemsCollection){
            context.putVar(var, obj);
            currentCell = cellRefGenerator.generateCellRef(index++, context);
            Size size = area.applyAt(currentCell, context);
            width = Math.max(width, size.getWidth());
            height = Math.max(height, size.getHeight());
            context.removeVar(var);
        }
        return new Size(width, height);
    }

    protected Collection calculateItemsCollection(Context context){
        ExpressionEvaluator expressionEvaluator = new JexlExpressionEvaluator(context.toMap());
        Object itemsObject = expressionEvaluator.evaluate(items);
        if( !(itemsObject instanceof Collection) ){
            throw new RuntimeException("items expression is not a collection");
        }
        return (Collection) itemsObject;
    }
}
