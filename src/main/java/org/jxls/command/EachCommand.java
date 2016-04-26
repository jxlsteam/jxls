package org.jxls.command;

import org.jxls.common.Size;
import org.jxls.area.Area;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.expression.JexlExpressionEvaluator;
import org.jxls.util.Util;

import java.util.Collection;

/**
 * Implements iteration over collection of items
 * 'items' is a bean name of the collection in context
 * 'var' is a name of a collection item to put into the context during the iteration
 * 'direction' defines expansion by rows (DOWN) or by columns (RIGHT). Default is DOWN.
 * 'cellRefGenerator' defines custom strategy for target cell references.
 * Date: Nov 10, 2009
 * @author Leonid Vysochyn
 */
public class EachCommand extends AbstractCommand {
    public enum Direction {RIGHT, DOWN}
    
    private String var;
    private String items;
    private String select;
    private Area area;
    private Direction direction = Direction.DOWN;
    private CellRefGenerator cellRefGenerator;

    public EachCommand() {
    }

    /**
     * @param var name of the key in the context to contain each collection items during iteration
     * @param items name of the collection bean in the context
     * @param direction defines processing by rows (DOWN - default) or columns (RIGHT)
     */
    public EachCommand(String var, String items, Direction direction) {
        this.var = var;
        this.items = items;
        this.direction = direction == null ? Direction.DOWN : direction;
    }

    public EachCommand(String var, String items, Area area) {
        this(var, items, area, Direction.DOWN);
    }
    
    public EachCommand(String var, String items, Area area, Direction direction) {
        this( var, items, direction );
        if( area != null ){
            this.area = area;
            addArea(this.area);
        }
    }

    /**
     *
     * @param var name of the key in the context to contain each collection items during iteration
     * @param items name of the collection bean in the context
     * @param area body area for this command
     * @param cellRefGenerator generates target cell ref for each collection item during iteration
     */
    public EachCommand(String var, String items, Area area, CellRefGenerator cellRefGenerator) {
        this(var, items, area, (Direction)null );
        this.cellRefGenerator = cellRefGenerator;
    }

    /**
     * Gets iteration directino
     * @return current direction for iteration
     */
    public Direction getDirection() {
        return direction;
    }

    /**
     * Sets iteration direction
     * @param direction
     */
    public void setDirection(Direction direction) {
        this.direction = direction;
    }

    public void setDirection(String direction){
        this.direction = Direction.valueOf(direction);
    }

    /**
     * Gets defined cell ref generator
     * @return current {@link CellRefGenerator} instance or null
     */
    public CellRefGenerator getCellRefGenerator() {
        return cellRefGenerator;
    }

    public void setCellRefGenerator(CellRefGenerator cellRefGenerator) {
        this.cellRefGenerator = cellRefGenerator;
    }

    public String getName() {
        return "each";
    }

    /**
     * Gets current variable name for collection item in the context during iteration
     * @return collection item key name in the context
     */
    public String getVar() {
        return var;
    }

    /**
     * Sets current variable name for collection item in the context during iteration
     * @param var
     */
    public void setVar(String var) {
        this.var = var;
    }

    /**
     * Gets collection bean name
     * @return collection bean name in the context
     */
    public String getItems() {
        return items;
    }

    /**
     * Sets collection bean name
     * @param items collection bean name in the context
     */
    public void setItems(String items) {
        this.items = items;
    }

    /**
     * Gets current 'select' expression for filtering out collection items
     * @return current 'select' expression or null if undefined
     */
    public String getSelect() {
        return select;
    }

    /**
     * Sets current 'select' expression for filtering collection
     * @param select filtering expression
     */
    public void setSelect(String select) {
        this.select = select;
    }

    @Override
    public Command addArea(Area area) {
        if( area == null ){
            return this;
        }
        if( areaList.size() >= 1){
            throw new IllegalArgumentException("You can add only a single area to 'each' command");
        }
        this.area = area;
        return super.addArea(area);
    }

    public Size applyAt(CellRef cellRef, Context context) {
        Collection itemsCollection = Util.transformToCollectionObject(getTransformationConfig().getExpressionEvaluator(), items, context);
        int width = 0;
        int height = 0;
        int index = 0;
        CellRef currentCell = cellRefGenerator != null ? cellRefGenerator.generateCellRef(index, context) : cellRef;
        JexlExpressionEvaluator selectEvaluator = null;
        if( select != null ){
            selectEvaluator = new JexlExpressionEvaluator( select );
        }
        for( Object obj : itemsCollection){
            context.putVar(var, obj);
            if( selectEvaluator != null && !Util.isConditionTrue(selectEvaluator, context) ){
                context.removeVar(var);
                continue;
            }
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

}
