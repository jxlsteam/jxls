package org.jxls.command;

import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import java.util.stream.Collectors;

import org.jxls.area.Area;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.Size;
import org.jxls.util.OrderByComparator;
import org.jxls.util.UtilWrapper;

/**
 * <p>Implements iteration over collection or array of items</p><ul>
 * <li>'items' is a bean name of the collection or array in context</li>
 * <li>'var' is a name of a collection item to put into the context during the iteration</li>
 * <li>'varIndex' is name of variable in context that holds current iteration index, 0 based. Use ${varIndex+1} for 1 based.</li>
 * <li>'direction' defines expansion by rows (DOWN) or by columns (RIGHT). Default is DOWN.</li>
 * <li>'select' holds an expression for filtering collection.</li>
 * <li>'groupBy' is the name for grouping (prepend var+".").</li>
 * <li>'groupOrder' defines the grouping order. Case does not matter.
 *     "ASC" for ascending, "DESC" for descending sort order. Other values or null: no sorting.</li>
 * <li>'orderBy' contains the names separated with comma and each with an optional postfix " ASC" (default) or " DESC" for the sort order. Prepend var+'.' before each name.</li>
 * <li>'multisheet' is the name of the sheet names container.</li>
 * <li>'cellRefGenerator' defines custom strategy for target cell references.</li>
 * </ul>
 * 
 * <p>The variables defined in 'var' and 'varIndex' will be saved using the special method Context.getRunVar()</p>
 *
 * @author Leonid Vysochyn
 */
public class EachCommand extends AbstractCommand {
    public static final String COMMAND_NAME = "each";
    static final String GROUP_DATA_KEY = "_group";
    /** Old behavior will be removed in a future release. */
    public static boolean oldSelectBehavior = false;

    private UtilWrapper util = new UtilWrapper();
    private String items;
    private String var;
    private String varIndex;
    public enum Direction {RIGHT, DOWN}
    private Direction direction = Direction.DOWN;
    private String select;
    private String groupBy;
    private String groupOrder;
    private String orderBy;
    private String multisheet;
    private CellRefGenerator cellRefGenerator;
    private Area area;
    private String useVarName;

    public EachCommand() {
    }

    /**
     * @param var       name of the key in the context to contain each collection items during iteration
     * @param items     name of collection or array in the context
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
     * @param items            name of collection or array in the context
     * @param area             body area for this command
     * @param cellRefGenerator generates target cell ref for each collection item during iteration
     */
    public EachCommand(String var, String items, Area area, CellRefGenerator cellRefGenerator) {
        this(var, items, area, (Direction) null);
        this.cellRefGenerator = cellRefGenerator;
    }

    @Override
    public String getName() {
        return COMMAND_NAME;
    }

    UtilWrapper getUtil() {
        return util;
    }

    void setUtil(UtilWrapper util) {
        this.util = util;
    }

    /**
     * Gets collection bean name
     *
     * @return collection name of collection or array in the context
     */
    public String getItems() {
        return items;
    }

    /**
     * Sets collection bean name
     *
     * @param items name of collection or array in the context
     */
    public void setItems(String items) {
        this.items = items;
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
     * @return variable name to put the current iteration index, 0 based
     */
    public String getVarIndex() {
        return varIndex;
    }

    public void setVarIndex(String varIndex) {
        this.varIndex = varIndex;
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
     * @return property name for grouping the collection
     */
    public String getGroupBy() {
        return groupBy;
    }

    /**
     * @param groupBy property name for grouping the collection.
     * You should write the run var name + "." before the property name.
     */
    public void setGroupBy(String groupBy) {
        this.groupBy = groupBy;
    }

    /**
     * @return group order
     */
    public String getGroupOrder() {
        return groupOrder;
    }

    /**
     * @param groupOrder group ordering: "ASC" for ascending, "DESC" for descending, other value or null: no sorting. Case does not matter.
     */
    public void setGroupOrder(String groupOrder) {
        this.groupOrder = groupOrder;
    }

    /**
     * @param orderBy property names for ordering the list.
     * You should write the run var name + "." before each property name.
     * You can write " ASC" or " DESC" after each property name for ascending/descending sorting order. ASC is the default.
     */
    public void setOrderBy(String orderBy) {
        this.orderBy = orderBy;
    }

    /**
     * @return property names for ordering the list
     */
    public String getOrderBy() {
        return orderBy;
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

    @Override
    public Size applyAt(CellRef cellRef, Context context) {
        Iterable<?> itemsCollection = prepareCollection(cellRef, context);
        Size size = processCollection(cellRef, context, itemsCollection);
        if (direction == Direction.DOWN) {
            getTransformer().adjustTableSize(cellRef, size);
        }
        return size;
    }
    
    protected Iterable<?> prepareCollection(CellRef cellRef, Context context) {
        Iterable<?> itemsCollection = null;
        try {
            itemsCollection = util.transformToIterableObject(getTransformationConfig().getExpressionEvaluator(), items, context);
        } catch (Exception e) {
            getTransformer().getExceptionHandler().handleEvaluationException(e, cellRef.toString(), items);
            itemsCollection = Collections.emptyList();
        }
        useVarName = var;
        boolean hasGroupBy = groupBy != null && !groupBy.isEmpty();
        if (hasGroupBy) {
            if (useVarName == null) {
                useVarName = GROUP_DATA_KEY;
            }
            if (select != null && !select.isEmpty() && !oldSelectBehavior) {
                CollectionFilter cf = new CollectionFilter(context, util, varIndex, null, select, null/*!!*/, null, null, null);
                cf.processCollection(itemsCollection, null, useVarName);
                itemsCollection = cf.getFilteredCollection();
            }
        }
        orderCollection(itemsCollection);
        if (hasGroupBy) {
            itemsCollection = util.groupIterable(itemsCollection, groupBy, groupOrder);
        }
        return itemsCollection;
    }
    
    @SuppressWarnings("unchecked")
    protected void orderCollection(Iterable<?> itemsCollection) {
        if (itemsCollection instanceof List && orderBy != null && !orderBy.trim().isEmpty()) {
            List<String> orderByProps = Arrays.asList(orderBy.split(","))
                    .stream().map(f -> removeVarPrefix(f.trim())).collect(Collectors.toList());
            OrderByComparator<Object> comp = new OrderByComparator<>(orderByProps, util);
            Collections.sort((List<Object>) itemsCollection, comp);
        }
    }
    
    protected String removeVarPrefix(String pVariable) {
        int o = pVariable.indexOf(".");
        if (o >= 0) {
            String pre = pVariable.substring(0, o).trim();
            if (pre.equals(var)) {
                return pVariable.substring(o + 1).trim();
            }
        }
        return pVariable;
    }

    protected Size processCollection(CellRef cellRef, Context context, Iterable<?> itemsCollection) {
        CollectionProcessor cp = new CollectionProcessor(context, util, varIndex, direction,
                groupBy == null || groupBy.isEmpty() || oldSelectBehavior ? select : null,
                groupBy, multisheet, cellRefGenerator, area);
        cp.initMultiSheet(cellRef, () -> getTransformationConfig());
        return cp.processCollection(itemsCollection, cellRef, useVarName);
    }
}
