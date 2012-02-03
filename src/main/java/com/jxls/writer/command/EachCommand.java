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
    public enum Direction {RIGHT, DOWN}
    
    String var;
    String items;
    boolean byRows = true;
    Command area;
    Direction direction = Direction.DOWN;

    public EachCommand(Size initialSize, String var, String items, Area area) {
        super(initialSize);
        this.var = var;
        this.items = items;
        this.area = area;
    }
    
    public EachCommand(Size initialSize, String var, String items, Area area, Direction direction) {
        super(initialSize);
        this.var = var;
        this.items = items;
        this.area = area;
        this.direction = direction == null ? Direction.DOWN : direction;
    }
    
    

    public Direction getDirection() {
        return direction;
    }

    public void setDirection(Direction direction) {
        this.direction = direction;
    }

    public Size applyAt(Cell cell, Context context) {
        Collection itemsCollection = calculateItemsCollection(context);
        Cell currentCell = cell;
        int width = 0;
        int height = 0;
        for( Object obj : itemsCollection){
            context.putVar(var, obj);
            Size size = area.applyAt(currentCell, context);
            if( direction == Direction.DOWN ){
                currentCell = new Cell(currentCell.getSheetIndex(), currentCell.getRow() + size.getHeight(), currentCell.getCol());
                width = Math.max(width, size.getWidth());
                height += size.getHeight();
            }else{
                currentCell = new Cell(currentCell.getSheetIndex(), currentCell.getRow(), currentCell.getCol() + size.getWidth());
                width += size.getWidth();
                height = Math.max( height, size.getHeight() );
            }
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
