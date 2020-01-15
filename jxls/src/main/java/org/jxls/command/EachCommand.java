package org.jxls.command;

import java.util.Arrays;
import java.util.Collection;
import java.util.Collections;
import java.util.List;

import org.jxls.area.Area;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.GroupData;
import org.jxls.common.JxlsException;
import org.jxls.common.Size;
import org.jxls.expression.ExpressionEvaluator;
import org.jxls.util.JxlsHelper;
import org.jxls.util.OrderByComparator;
import org.jxls.util.UtilWrapper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * <p>Implements iteration over collection of items</p><ul>
 * <li>'items' is a bean name of the collection in context</li>
 * <li>'var' is a name of a collection item to put into the context during the iteration</li>
 * <li>'direction' defines expansion by rows (DOWN) or by columns (RIGHT). Default is DOWN.</li>
 * <li>'cellRefGenerator' defines custom strategy for target cell references.</li>
 * <li>'select' holds an expression for filtering collection.</li>
 * <li>'multisheet' is the name of the sheet names container.</li>
 * <li>'groupBy' is the name for grouping.</li>
 * <li>'groupOrder' defines the grouping order. Case does not matter.
 *     "asc" for ascending, "desc" for descending sort order. Other values or null: no sorting.</li>
 * </ul>
 *
 * @author Leonid Vysochyn
 */
public class EachCommand extends AbstractCommand {
    public static final String COMMAND_NAME = "each";
    private static Logger logger = LoggerFactory.getLogger(EachCommand.class);
    static final String GROUP_DATA_KEY = "_group";

    private UtilWrapper util = new UtilWrapper();
    private Area area;
    private String items;
    private String var;
    public enum Direction {RIGHT, DOWN}
    private Direction direction = Direction.DOWN;
    private CellRefGenerator cellRefGenerator;
    private String select;
    private String multisheet;
    private String groupBy;
    private String groupOrder;
    private String varIndex;
    private String orderBy;

    public EachCommand() {
    }

    /**
     * @param var       name of the key in the context to contain each collection items during iteration
     * @param items     name of the collection bean in the context
     * @param direction defines processing by rows (DOWN - default) or columns (RIGHT)
     */
    public EachCommand(String var, String items, Direction direction) {
        this.var = var;
        this.items = items;
        this.direction = direction == null ? Direction.DOWN : direction;
    }

    public EachCommand(String items, Area area) {
        this(null, items, area);
    }

    public EachCommand(String var, String items, Area area) {
        this(var, items, area, Direction.DOWN);
    }

    public EachCommand(String var, String items, Area area, Direction direction) {
        this(var, items, direction);
        if (area != null) {
            this.area = area;
            addArea(this.area);
        }
    }

    /**
     * @param var              name of the key in the context to contain each collection items during iteration
     * @param items            name of the collection bean in the context
     * @param area             body area for this command
     * @param cellRefGenerator generates target cell ref for each collection item during iteration
     */
    public EachCommand(String var, String items, Area area, CellRefGenerator cellRefGenerator) {
        this(var, items, area, (Direction) null);
        this.cellRefGenerator = cellRefGenerator;
    }

    /**
     * @return variable name to put the current iteration index
     */
    public String getVarIndex() {
        return varIndex;
    }

    public void setVarIndex(String varIndex) {
        this.varIndex = varIndex;
    }

    UtilWrapper getUtil() {
        return util;
    }

    void setUtil(UtilWrapper util) {
        this.util = util;
    }
    
    /**
     * Gets iteration direction
     *
     * @return current direction for iteration
     */
    public Direction getDirection() {
        return direction;
    }

    /**
     * Sets iteration direction
     *
     * @param direction iteration direction
     */
    public void setDirection(Direction direction) {
        this.direction = direction;
    }

    /**
     * @param direction "DOWN" or "RIGHT"
     */
    public void setDirection(String direction) {
        this.direction = Direction.valueOf(direction);
    }

    /**
     * Gets defined cell ref generator
     *
     * @return current {@link CellRefGenerator} instance or null
     */
    public CellRefGenerator getCellRefGenerator() {
        return cellRefGenerator;
    }

    public void setCellRefGenerator(CellRefGenerator cellRefGenerator) {
        this.cellRefGenerator = cellRefGenerator;
    }

    @Override
    public String getName() {
        return COMMAND_NAME;
    }

    /**
     * Gets current variable name for collection item in the context during iteration
     *
     * @return collection item key name in the context
     */
    public String getVar() {
        return var;
    }

    /**
     * Sets current variable name for collection item in the context during iteration
     *
     * @param var name of the loop var
     */
    public void setVar(String var) {
        this.var = var;
    }

    /**
     * Gets collection bean name
     *
     * @return collection bean name in the context
     */
    public String getItems() {
        return items;
    }

    /**
     * Sets collection bean name
     *
     * @param items collection bean name in the context
     */
    public void setItems(String items) {
        this.items = items;
    }

    /**
     * Gets current 'select' expression for filtering out collection items
     *
     * @return current 'select' expression or null if undefined
     */
    public String getSelect() {
        return select;
    }

    /**
     * Sets current 'select' expression for filtering collection
     *
     * @param select filtering expression
     */
    public void setSelect(String select) {
        this.select = select;
    }

    /**
     * @return Context variable name holding a list of Excel sheet names to output the collection to
     */
    public String getMultisheet() {
        return multisheet;
    }

    /**
     * Sets name of context variable holding a list of Excel sheet names to output the collection to
     *
     * @param multisheet var name
     */
    public void setMultisheet(String multisheet) {
        this.multisheet = multisheet;
    }

    /**
     * @return property name for grouping the collection
     */
    public String getGroupBy() {
        return groupBy;
    }

    /**
     * @param groupBy property name for grouping the collection
     */
    public void setGroupBy(String groupBy) {
        this.groupBy = groupBy;
    }

    /**
     * @param orderBy property name for ordering the collection
     */
    public void setOrderBy(String orderBy) {
        this.orderBy = orderBy;
    }

    /**
     * @return group order
     */
    public String getGroupOrder() {
        return groupOrder;
    }

    /**
     * @param groupOrder group ordering: asc = ascending, desc = descending, other value or null: no sorting. Case does not matter.
     */
    public void setGroupOrder(String groupOrder) {
        this.groupOrder = groupOrder;
    }

    @Override
    public Command addArea(Area area) {
        if (area == null) {
            return this;
        }
        if (areaList.size() >= 1) {
            throw new IllegalArgumentException("You can add only a single area to 'each' command");
        }
        this.area = area;
        return super.addArea(area);
    }

    @SuppressWarnings("unchecked")
    @Override
    public Size applyAt(CellRef cellRef, Context context) {
        Iterable<?> itemsCollection = null;
        try {
            itemsCollection = util.transformToIterableObject(getTransformationConfig().getExpressionEvaluator(), items, context);
            if (itemsCollection instanceof List) {
                OrderByComparator<Object> comp = null;
                if (orderBy != null && !orderBy.isEmpty()) {
                    List<String> orderByProps = Arrays.asList(orderBy.split(","));
                    comp = new OrderByComparator<>(orderByProps, util);
                    Collections.sort((List<Object>) itemsCollection, comp);
                }
            }
        } catch (Exception e) {
            logger.warn("Failed to evaluate collection expression {}", items, e);
            itemsCollection = Collections.emptyList();
        }
        Size size;
        if (groupBy == null || groupBy.length() == 0) {
            size = processCollection(context, itemsCollection, cellRef, var);
        } else {
            Collection<GroupData> groupedData = util.groupIterable(itemsCollection, groupBy, groupOrder);
            String groupVar = var != null ? var : GROUP_DATA_KEY;
            size = processCollection(context, groupedData, cellRef, groupVar);
        }
        if (direction == Direction.DOWN) {
            getTransformer().adjustTableSize(cellRef, size);
        }
        return size;
    }

    private Size processCollection(Context context, Iterable<?> itemsCollection, CellRef cellRef, String varName) {
        int index = 0;
        int newWidth = 0;
        int newHeight = 0;

        CellRefGenerator cellRefGenerator = this.cellRefGenerator;
        if (cellRefGenerator == null && multisheet != null) {
            List<String> sheetNameList = extractSheetNameList(context);
            cellRefGenerator = sheetNameList == null
                    ? new DynamicSheetNameGenerator(multisheet, cellRef, getTransformationConfig().getExpressionEvaluator())
                    : new SheetNameGenerator(sheetNameList, cellRef);
        }
        
        ExpressionEvaluator selectEvaluator = null;
        if (select != null) {
            selectEvaluator = JxlsHelper.getInstance().createExpressionEvaluator(select);
        }

        CellRef currentCell = cellRef;
        Object currentVarObject = context.getVar(varName);
        int currentIndex = 0;
        for (Object obj : itemsCollection) {
            context.putVar(varName, obj);
            if (varIndex != null ){
                context.putVar(varIndex, currentIndex);
            }
            if (selectEvaluator != null && !util.isConditionTrue(selectEvaluator, context)) {
                context.removeVar(varName);
                continue;
            }
            if (cellRefGenerator != null) {
                currentCell = cellRefGenerator.generateCellRef(index++, context);
            }
            if (currentCell == null) {
                break;
            }
            Size size = area.applyAt(currentCell, context);
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
            currentIndex++;
        }
        restoreVarObject(context, varName, currentVarObject);
        return new Size(newWidth, newHeight);
    }

    private void restoreVarObject(Context context, String varName, Object varObject) {
        if (varObject != null) {
            context.putVar(varName, varObject);
        } else {
            context.removeVar(varName);
        }
    }

    @SuppressWarnings("unchecked")
    private List<String> extractSheetNameList(Context context) {
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
}
