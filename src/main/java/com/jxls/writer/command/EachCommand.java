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

    public EachCommand(Pos pos, Size initialSize, String var, String items, Command area) {
        super(pos, initialSize);
        this.var = var;
        this.items = items;
        this.area = area;
    }

    public Size applyAt(Pos pos, Context context) {
        Collection itemsCollection = calculateItemsCollection(context);
        Pos currentPos = pos;
        int width = 0;
        int height = 0;
        for( Object obj : itemsCollection){
            context.putVar(var, obj);
            Size size = area.applyAt(currentPos, context);
            if( byRows ){
                currentPos = new Pos(currentPos.getRow() + size.getHeight(), currentPos.getCol());
                width = Math.max(width, size.getWidth());
                height += size.getHeight();
            }else{
                currentPos = new Pos(currentPos.getRow(), currentPos.getCol() + size.getWidth());
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
