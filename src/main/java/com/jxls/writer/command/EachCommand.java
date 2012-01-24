package com.jxls.writer.command;

import com.jxls.writer.*;
import com.jxls.writer.expression.ExpressionEvaluator;
import com.jxls.writer.expression.JexlExpressionEvaluator;

import java.util.Collection;

/**
 * Date: Nov 10, 2009
 *
 * @author Leonid Vysochyn
 */
public class EachCommand extends AbstractCommand {
    String var;
    String items;
    boolean byRows = true;
    Command area;

    public EachCommand(Cell cell, Size initialSize, String var, String items, Command area) {
        super(cell, initialSize);
        this.var = var;
        this.items = items;
        this.area = area;
    }

    public Size applyAt(Cell cell, Context context) {
        Collection itemsCollection = calculateItemsCollection(context);
        Cell currentCell = cell;
        int width = 0;
        int height = 0;
        for( Object obj : itemsCollection){
            context.putVar(var, obj);
            Size size = area.applyAt(currentCell, context);
            if( byRows ){
                currentCell = new Cell(currentCell.getCol() + size.getHeight(), currentCell.getRow());
                width = Math.max(width, size.getWidth());
                height += size.getHeight();
            }else{
                currentCell = new Cell(currentCell.getCol(), currentCell.getRow() + size.getWidth());
                width += size.getWidth();
                height = Math.max( height, size.getHeight() );
            }
            context.removeVar(var);
        }
        return new Size(width, height);
    }

    public Size getSize(Context context) {
        Collection itemsCollection = calculateItemsCollection(context);
        int width = 0;
        int height = 0;
        for( Object obj : itemsCollection){
            context.putVar(var, obj);
            Size size = area.getSize(context);
            if( byRows ){
                width = Math.max(width, size.getWidth());
                height += size.getHeight();
            }else{
                width += size.getWidth();
                height = Math.max( height, size.getHeight() );
            }
            context.removeVar( var );
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
