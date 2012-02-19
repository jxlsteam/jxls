package com.jxls.writer.command;

import com.jxls.writer.area.Area;
import com.jxls.writer.common.CellRef;
import com.jxls.writer.common.Context;
import com.jxls.writer.common.Size;
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
    Area area;
    Direction direction = Direction.DOWN;
    CellRefGenerator cellRefGenerator;

    public EachCommand(String var, String items, Direction direction) {
        this.var = var;
        this.items = items;
        this.direction = direction;
    }

    public EachCommand(String var, String items, Area area) {
        this(var, items, area, Direction.DOWN);
    }
    
    public EachCommand(String var, String items, Area area, Direction direction) {
        this.var = var;
        this.items = items;
        this.area = area;
        this.direction = direction == null ? Direction.DOWN : direction;
        addArea(this.area);
    }

    public EachCommand(String var, String items, Area area, CellRefGenerator cellRefGenerator) {
        this.var = var;
        this.items = items;
        this.area = area;
        this.cellRefGenerator = cellRefGenerator;
        addArea(area);
    }

    public Direction getDirection() {
        return direction;
    }

    public void setDirection(Direction direction) {
        this.direction = direction;
    }

    public CellRefGenerator getCellRefGenerator() {
        return cellRefGenerator;
    }

    public void setCellRefGenerator(CellRefGenerator cellRefGenerator) {
        this.cellRefGenerator = cellRefGenerator;
    }

    public String getName() {
        return "each";
    }

    public String getVar() {
        return var;
    }

    public String getItems() {
        return items;
    }

    @Override
    public void addArea(Area area) {
        if( areaList.size() >= 1){
            throw new IllegalArgumentException("You can add only a single area to 'Each' command");
        }
        super.addArea(area);    
        this.area = area;
    }

    public Size applyAt(CellRef cellRef, Context context) {
        Collection itemsCollection = calculateItemsCollection(context);
        int width = 0;
        int height = 0;
        int index = 0;
        CellRef currentCell = cellRefGenerator != null ? cellRefGenerator.generateCellRef(index, context) : cellRef;
        for( Object obj : itemsCollection){
            context.putVar(var, obj);
            Size size = area.applyAt(currentCell, context);
            index++;
            if( cellRefGenerator != null ){
                width = Math.max(width, size.getWidth());
                height = Math.max(height, size.getHeight());
                currentCell = cellRefGenerator.generateCellRef(index, context);
            }else if( direction == Direction.DOWN ){
                currentCell = new CellRef(currentCell.getSheetName(), currentCell.getRow() + size.getHeight(), currentCell.getCol());
                width = Math.max(width, size.getWidth());
                height += size.getHeight();
            }else{
                currentCell = new CellRef(currentCell.getSheetName(), currentCell.getRow(), currentCell.getCol() + size.getWidth());
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
