package org.jxls.command;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.Collections;
import java.util.Iterator;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;
import java.util.TreeSet;
import java.util.stream.Collectors;

import org.jxls.area.Area;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.GroupData;
import org.jxls.common.JxlsException;
import org.jxls.common.ObjectPropertyAccess;
import org.jxls.common.Size;
import org.jxls.expression.ExpressionEvaluator;
import org.jxls.logging.JxlsLogger;
import org.jxls.util.OrderByComparator;

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

    private String items;
    private String var;
    private String varIndex;
    public enum Direction {
        /** Dynamic columns. See also jx:params(formulaStrategy="BY_COLUMN") */
        RIGHT,
        /** Dynamic rows (default) */
        DOWN;
    }
    private Direction direction = Direction.DOWN;
    private String select;
    private String groupBy;
    private String groupOrder;
    private String orderBy;
    private String multisheet;
    private CellRefGenerator cellRefGenerator;
    private Area area;
    /** Old behavior will be removed in a future release. */
    private boolean oldSelectBehavior = false;

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

    public boolean isOldSelectBehavior() {
        return oldSelectBehavior;
    }

    public void setOldSelectBehavior(boolean oldSelectBehavior) {
        this.oldSelectBehavior = oldSelectBehavior;
    }
    
    public void setOldSelectBehavior(String oldSelectBehavior) {
        this.oldSelectBehavior = "true".equalsIgnoreCase(oldSelectBehavior);
    }

    @Override
    public Command addArea(Area area) {
        if (area == null) {
            return this;
        }
        if (areaList.size() >= 1) {
            throw new JxlsException("You can add only a single area to 'each' command");
        }
        this.area = area;
        return super.addArea(area);
    }

    @Override
    public Size applyAt(CellRef cellRef, Context context) {
        Iterable<?> itemsCollection = null;
        try {
            itemsCollection = transformToIterableObject(items, context);
            orderCollection(itemsCollection);
        } catch (Exception e) {
            getLogger().handleEvaluationException(e, cellRef.toString(), items);
            itemsCollection = Collections.emptyList();
        }
        Size size;
        if (groupBy == null || groupBy.length() == 0) {
            size = processCollection(context, itemsCollection, cellRef, var, select);
        } else {
            String selectExpression = select;
            if (selectExpression != null && !selectExpression.isEmpty() // filtering needed?
                    && !oldSelectBehavior) { // new behavior
                itemsCollection = filter(context, itemsCollection, selectExpression);
                selectExpression = null;
            }
            Collection<GroupData> groupedData = groupIterable(itemsCollection, groupBy, groupOrder, var, getLogger());
            String groupVar = var != null ? var : GROUP_DATA_KEY;
            size = processCollection(context, groupedData, cellRef, groupVar, selectExpression);
        }
        if (direction == Direction.DOWN) {
            getTransformer().adjustTableSize(cellRef, size);
        }
        return size;
    }
    
    private void orderCollection(Iterable<?> itemsCollection) {
        if (itemsCollection instanceof List<?> itemsList && orderBy != null && !orderBy.trim().isEmpty()) {
            List<String> orderByProps = Arrays.asList(orderBy.split(","))
                    .stream().map(f -> removeVarPrefix(f, var)).collect(Collectors.toList());
            itemsList.sort(new OrderByComparator<>(orderByProps));
        }
    }

    private Iterable<?> filter(Context context, Iterable<?> itemsCollection, String selectExpression) {
        List<Object> filteredList = new ArrayList<>();
        ExpressionEvaluator selectEvaluator = context.getExpressionEvaluator(selectExpression);
        int currentIndex = 0;
        try (RunVar runVar = new RunVar(var, varIndex, context)) {
            for (Object obj : itemsCollection) {
                runVar.put(obj, Integer.valueOf(currentIndex));
                if (selectEvaluator.isConditionTrue(context)) {
                    filteredList.add(obj);
                }
                currentIndex++;
            }
        }
        return filteredList;
    }

    private Size processCollection(Context context, Iterable<?> itemsCollection, CellRef cellRef, String varName, String selectExpression) {
        int index = 0;
        int newWidth = 0;
        int newHeight = 0;

        CellRefGenerator cellRefGenerator = this.cellRefGenerator;
        if (cellRefGenerator == null && multisheet != null) {
            List<String> sheetNameList = extractSheetNameList(context);
            cellRefGenerator = sheetNameList == null
                    ? new DynamicSheetNameGenerator(multisheet, cellRef)
                    : new SheetNameGenerator(sheetNameList, cellRef);
        }
        
        ExpressionEvaluator selectEvaluator = null;
        if (selectExpression != null) {
            selectEvaluator = context.getExpressionEvaluator(selectExpression);
        }

        CellRef currentCell = cellRef;
        int currentIndex = 0;
        try (RunVar runVar = new RunVar(varName, varIndex, context)) {
            for (Object obj : itemsCollection) {
                runVar.put(obj, Integer.valueOf(currentIndex));
                if (selectEvaluator != null && !selectEvaluator.isConditionTrue(context)) {
                    continue;
                }
                if (cellRefGenerator != null) {
                    currentCell = cellRefGenerator.generateCellRef(index++, context, getLogger());
                }
                if (currentCell == null) {
                    break;
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
                currentIndex++;
            }
        }
        return new Size(newWidth, newHeight);
    }
    
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
    
    private static String removeVarPrefix(String pVariable, String pVar) {
    	pVariable = pVariable.trim();
        int o = pVariable.indexOf(".");
        if (o >= 0) {
            String pre = pVariable.substring(0, o).trim();
            if (pre.equals(pVar)) {
                return pVariable.substring(o + 1).trim();
            }
        }
        return pVariable;
    }
    
    /**
     * Groups items from an iterable collection using passed group property and group order
     * @param iterable iterable object
     * @param groupProperty property to use to group the items
     * @param groupOrder an order to sort the groups
     * @return a collection of group data objects
     */
    static Collection<GroupData> groupIterable(Iterable<?> iterable, String groupProperty, String groupOrder, String var, JxlsLogger logger) {
        Collection<GroupData> result = new ArrayList<>();
        if (iterable == null) {
            return result;
        }
        Set<Object> groupByValues;
        if (groupOrder != null) {
            if ("desc".equalsIgnoreCase(groupOrder) || "desc_ignoreCase".equalsIgnoreCase(groupOrder)) {
                groupByValues = new TreeSet<>(Collections.reverseOrder());
            } else {
                groupByValues = new TreeSet<>();
            }
        } else {
            groupByValues = new LinkedHashSet<>();
        }
        boolean ignoreCase = groupOrder != null && groupOrder.toLowerCase().endsWith("_ignorecase");
        for (Object bean : iterable) {
            Object groupKey = getGroupKey(bean, groupProperty, var, logger);
            if (ignoreCase && groupKey instanceof String s) {
                groupKey = s.toLowerCase();
            }
            groupByValues.add(groupKey);
        }
        for (Iterator<Object> iterator = groupByValues.iterator(); iterator.hasNext();) {
            Object groupValue = iterator.next();
            List<Object> groupItems = new ArrayList<>();
            for (Object bean : iterable) {
                Object groupKey = getGroupKey(bean, groupProperty, var, logger);
                boolean eq;
                if (ignoreCase && groupValue instanceof String a && groupKey instanceof String b) {
                    eq = a.equalsIgnoreCase(b);
                } else {
                    eq = groupValue.equals(groupKey);
                }
                if (eq) {
                    groupItems.add(bean);
                }
            }
            if (!groupItems.isEmpty()) {
                result.add(new GroupData(groupItems.get(0), groupItems));
            }
        }
        return result;
    }

    private static Object getGroupKey(Object bean, String propertyName, String var, JxlsLogger logger) {
        Object ret = ObjectPropertyAccess.getObjectProperty(bean, removeVarPrefix(propertyName, var), logger);
        return ret == null ? "null" : ret;
    }
}
