package org.jxls.util;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.Iterator;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeSet;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.jxls.area.Area;
import org.jxls.command.Command;
import org.jxls.command.EachCommand;
import org.jxls.common.CellRef;
import org.jxls.common.CellRefColPrecedenceComparator;
import org.jxls.common.CellRefRowPrecedenceComparator;
import org.jxls.common.Context;
import org.jxls.common.GroupData;
import org.jxls.common.JxlsException;
import org.jxls.expression.EvaluationException;
import org.jxls.expression.ExpressionEvaluator;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Utility class with various helper methods used by other classes
 *
 * @author Leonid Vysochyn
 */
public class Util {
    private static Logger logger = LoggerFactory.getLogger(Util.class);
    public static final String regexJointedLookBehind = "(?<!U_\\([^)]{0,100})";
    public static final String regexSimpleCellRef = "[a-zA-Z]+[0-9]+";
    public static final String regexCellRef = "([a-zA-Z_]+[a-zA-Z0-9_]*![a-zA-Z]+[0-9]+|(?<!\\d)[a-zA-Z]+[0-9]+|'[^?\\\\/:'*]+'![a-zA-Z]+[0-9]+)";
    public static final String regexAreaRef = regexCellRef + ":" + regexSimpleCellRef;
    public static final Pattern regexAreaRefPattern = Pattern.compile(regexAreaRef);
    public static final String regexCellRefExcludingJointed = regexJointedLookBehind + regexCellRef;
    private static final Pattern regexCellRefExcludingJointedPattern = Pattern.compile(regexCellRefExcludingJointed);
    private static final Pattern regexCellRefPattern = Pattern.compile(regexCellRef);
    public static final String regexJointedCellRef = "U_\\([^\\)]+\\)";
    private static final Pattern regexJointedCellRefPattern = Pattern.compile(regexJointedCellRef);
    public static final String regexExcludePrefixSymbols = "(?<!\\w)";

    /**
     * Parses a formula and returns a list of cell names used in it
     * E.g. for formula "B4*(1+C4)" the returned list will contain "B4", "C4"
     * @param formula string
     * @return a list of cell names used in the formula
     */
    public static List<String> getFormulaCellRefs(String formula) {
        return getStringPartsByPattern(formula, regexCellRefExcludingJointedPattern);
    }

    private static List<String> getStringPartsByPattern(String str, Pattern pattern) {
        List<String> cellRefs = new ArrayList<String>();
        if (str != null) {
            Matcher cellRefMatcher = pattern.matcher(str);
            while (cellRefMatcher.find()) {
                cellRefs.add(cellRefMatcher.group());
            }
        }
        return cellRefs;
    }

    /**
     * Parses a formula to extract a list of so called "jointed cells"
     * The jointed cells are cells combined with a special notation "U_(cell1, cell2)" into a single cell
     * They are used in formulas like this "$[SUM(U_(F8,F13))]".
     * Here the formula will use both F8 and F13 source cells to calculate the sum
     * @param formula a formula string to parse
     * @return a list of jointed cells used in the formula
     */
    public static List<String> getJointedCellRefs(String formula) {
        return getStringPartsByPattern(formula, regexJointedCellRefPattern);
    }

    /**
     * Parses a "jointed cell" reference and extracts individual cell references
     * @param jointedCellRef a jointed cell reference to parse
     * @return a list of cell names extracted from the jointed cell reference
     */
    public static List<String> getCellRefsFromJointedCellRef(String jointedCellRef) {
        return getStringPartsByPattern(jointedCellRef, regexCellRefPattern);
    }

    /**
     * Checks if the formula contains jointed cell references
     * Jointed references have format U_(cell1, cell2) e.g. $[SUM(U_(F8,F13))]
     * @param formula string
     * @return true if the formula contains jointed cell references
     */
    public static boolean formulaContainsJointedCellRef(String formula) {
        return regexJointedCellRefPattern.matcher(formula).find();
    }

    /**
     * Combines a list of cell references into a range
     * E.g. for cell references A1, A2, A3, A4 it returns A1:A4
     * @param targetCellDataList
     * @return a range containing all the cell references if such range exists or otherwise the passed cells separated by commas
     */
    public static String createTargetCellRef(List<CellRef> targetCellDataList) {
        String resultRef = "";
        if (targetCellDataList == null || targetCellDataList.isEmpty()) {
            return resultRef;
        }
        List<String> cellRefs = new ArrayList<String>();
        boolean rowRange = true;
        boolean colRange = true;
        Iterator<CellRef> iterator = targetCellDataList.iterator();
        CellRef firstCellRef = iterator.next();
        cellRefs.add(firstCellRef.getCellName());
        String sheetName = firstCellRef.getSheetName();
        int row = firstCellRef.getRow();
        int col = firstCellRef.getCol();
        while (iterator.hasNext()) {
            CellRef cellRef = iterator.next();
            if ((rowRange || colRange) && !cellRef.getSheetName().equals(sheetName)) {
                rowRange = false;
                colRange = false;
            }
            if (rowRange && !(cellRef.getRow() - row == 1 && cellRef.getCol() == col)) {
                rowRange = false;
            }
            if (colRange && !(cellRef.getCol() - col == 1 && cellRef.getRow() == row)) {
                colRange = false;
            }
            sheetName = cellRef.getSheetName();
            row = cellRef.getRow();
            col = cellRef.getCol();
            cellRefs.add(cellRef.getCellName());
        }
        if ((rowRange || colRange) && cellRefs.size() > 1) {
            resultRef = cellRefs.get(0) + ":" + cellRefs.get(cellRefs.size() - 1);
        } else {
            resultRef = joinStrings(cellRefs, ",");
        }
        return resultRef;
    }

    /**
     * Joins strings with a separator
     * @param strings
     * @param separator
     * @return a string consisting of all the passed strings joined with the separator
     */
    public static String joinStrings(List<String> strings, String separator) {
        StringBuilder sb = new StringBuilder();
        String sep = "";
        for (String s : strings) {
            sb.append(sep).append(s);
            sep = separator;
        }
        return sb.toString();
    }

    /**
     * Groups a list of cell references into a list ranges which can be used in a formula substitution
     * @param cellRefList a list of cell references
     * @param targetRangeCount a number of ranges to use when grouping
     * @return a list of cell ranges grouped by row or by column
     */
    public static List<List<CellRef>> groupByRanges(List<CellRef> cellRefList, int targetRangeCount) {
        List<List<CellRef>> colRanges = groupByColRange(cellRefList);
        if (targetRangeCount == 0 || colRanges.size() == targetRangeCount) {
            return colRanges;
        }
        List<List<CellRef>> rowRanges = groupByRowRange(cellRefList);
        if (rowRanges.size() == targetRangeCount) {
            return rowRanges;
        } else {
            return colRanges;
        }
    }

    /**
     * Groups a list of cell references in a column into a list of ranges
     * @param cellRefList
     * @return a list of cell reference groups
     */
    public static List<List<CellRef>> groupByColRange(List<CellRef> cellRefList) {
        List<List<CellRef>> rangeList = new ArrayList<List<CellRef>>();
        if (cellRefList == null || cellRefList.size() == 0) {
            return rangeList;
        }
        List<CellRef> cellRefListCopy = new ArrayList<CellRef>(cellRefList);
        Collections.sort(cellRefListCopy, new CellRefColPrecedenceComparator());

        String sheetName = cellRefListCopy.get(0).getSheetName();
        int row = cellRefListCopy.get(0).getRow();
        int col = cellRefListCopy.get(0).getCol();
        List<CellRef> currentRange = new ArrayList<CellRef>();
        currentRange.add(cellRefListCopy.get(0));
        boolean rangeComplete = false;
        for (int i = 1; i < cellRefListCopy.size(); i++) {
            CellRef cellRef = cellRefListCopy.get(i);
            if (!cellRef.getSheetName().equals(sheetName)) {
                rangeComplete = true;
            } else {
                int rowDelta = cellRef.getRow() - row;
                int colDelta = cellRef.getCol() - col;
                if (rowDelta == 1 && colDelta == 0) {
                    currentRange.add(cellRef);
                } else {
                    rangeComplete = true;
                }
            }
            sheetName = cellRef.getSheetName();
            row = cellRef.getRow();
            col = cellRef.getCol();
            if (rangeComplete) {
                rangeList.add(currentRange);
                currentRange = new ArrayList<CellRef>();
                currentRange.add(cellRef);
                rangeComplete = false;
            }
        }
        rangeList.add(currentRange);
        return rangeList;
    }

    /**
     * Groups a list of cell references in a row into a list of ranges
     * @param cellRefList
     * @return
     */
    public static List<List<CellRef>> groupByRowRange(List<CellRef> cellRefList) {
        List<List<CellRef>> rangeList = new ArrayList<List<CellRef>>();
        if (cellRefList == null || cellRefList.size() == 0) {
            return rangeList;
        }
        List<CellRef> cellRefListCopy = new ArrayList<CellRef>(cellRefList);
        Collections.sort(cellRefListCopy, new CellRefRowPrecedenceComparator());

        String sheetName = cellRefListCopy.get(0).getSheetName();
        int row = cellRefListCopy.get(0).getRow();
        int col = cellRefListCopy.get(0).getCol();
        List<CellRef> currentRange = new ArrayList<CellRef>();
        currentRange.add(cellRefListCopy.get(0));
        boolean rangeComplete = false;
        for (int i = 1; i < cellRefListCopy.size(); i++) {
            CellRef cellRef = cellRefListCopy.get(i);
            if (!cellRef.getSheetName().equals(sheetName)) {
                rangeComplete = true;
            } else {
                int rowDelta = cellRef.getRow() - row;
                int colDelta = cellRef.getCol() - col;
                if (colDelta == 1 && rowDelta == 0) {
                    currentRange.add(cellRef);
                } else {
                    rangeComplete = true;
                }
            }
            sheetName = cellRef.getSheetName();
            row = cellRef.getRow();
            col = cellRef.getCol();
            if (rangeComplete) {
                rangeList.add(currentRange);
                currentRange = new ArrayList<CellRef>();
                currentRange.add(cellRef);
                rangeComplete = false;
            }
        }
        rangeList.add(currentRange);
        return rangeList;
    }

    /**
     * Evaluates if the passed condition is true
     * @param evaluator expression evaluator instance
     * @param condition condition string
     * @param context Jxls context to use for evaluation
     * @return true if the condition is evaluated to true or false otherwise
     */
    public static Boolean isConditionTrue(ExpressionEvaluator evaluator, String condition, Context context) {
        Object conditionResult = evaluator.evaluate(condition, context.toMap());
        if (!(conditionResult instanceof Boolean)) {
            throw new JxlsException("Condition result is not a boolean value - " + condition);
        }
        return (Boolean) conditionResult;
    }

    /**
     * Evaluates if the passed condition is true
     * @param evaluator
     * @param context Jxls context to use for evaluation
     * @return true if the condition is evaluated to true or false otherwise
     */
    public static Boolean isConditionTrue(ExpressionEvaluator evaluator, Context context) {
        Object conditionResult = evaluator.evaluate(context.toMap());
        if (!(conditionResult instanceof Boolean)) {
            throw new EvaluationException("Condition result is not a boolean value - " + evaluator.getExpression());
        }
        return (Boolean) conditionResult;
    }

    /**
     * Dynamically sets an object property via reflection
     * @param obj
     * @param propertyName
     * @param propertyValue
     * @param ignoreNonExisting
     */
    public static void setObjectProperty(Object obj, String propertyName, String propertyValue, boolean ignoreNonExisting) {
        try {
            setObjectProperty(obj, propertyName, propertyValue);
        } catch (Exception e) {
            String msg = "failed to set property '" + propertyName + "' to value '" + propertyValue + "' for object " + obj;
            if (ignoreNonExisting) {
                logger.info(msg, e);
            } else {
                logger.warn(msg);
                throw new IllegalArgumentException(e);
            }
        }
    }

    /**
     * Dynamically sets an object property via reflection
     * @param obj
     * @param propertyName
     * @param propertyValue
     * @throws NoSuchMethodException
     * @throws InvocationTargetException
     * @throws IllegalAccessException
     */
    public static void setObjectProperty(Object obj, String propertyName, String propertyValue)
            throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {
        String name = "set" + Character.toUpperCase(propertyName.charAt(0)) + propertyName.substring(1);
        Method method = obj.getClass().getMethod(name, new Class[] { String.class });
        method.invoke(obj, propertyValue);
    }

    /**
     * Gets value of the passed object property name
     * @param obj
     * @param propertyName
     * @param failSilently
     * @return
     */
    public static Object getObjectProperty(Object obj, String propertyName, boolean failSilently) {
        try {
            return getObjectProperty(obj, propertyName);
        } catch (Exception e) {
            String msg = "failed to get property '" + propertyName + "' of object " + obj;
            if (failSilently) {
                logger.info(msg, e);
                return null;
            } else {
                logger.warn(msg);
                throw new IllegalArgumentException(e);
            }
        }
    }

    /**
     * Gets value of the passed object property name
     * @param obj
     * @param propertyName
     * @return
     * @throws NoSuchMethodException
     * @throws InvocationTargetException
     * @throws IllegalAccessException
     */
    public static Object getObjectProperty(Object obj, String propertyName) throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {
        if (obj instanceof Map) {
            return ((Map<?, ?>) obj).get(propertyName);
        }
        String name = "get" + Character.toUpperCase(propertyName.charAt(0)) + propertyName.substring(1);
        Method method = obj.getClass().getMethod(name);
        return method.invoke(obj);
    }

    /**
     * Similar to {@link this#groupIterable(Iterable, String, String)} method but works for collections
     * @param collection
     * @param groupProperty
     * @param groupOrder
     * @return a collection of group data objects
     */
    public static Collection<GroupData> groupCollection(Collection<?> collection, String groupProperty, String groupOrder) {
        Collection<GroupData> result = new ArrayList<GroupData>();
        if (collection == null) {
            return result;
        }
        Set<Object> groupByValues;
        if (groupOrder != null) {
            if ("desc".equalsIgnoreCase(groupOrder)) {
                groupByValues = new TreeSet<>(Collections.reverseOrder());
            } else {
                groupByValues = new TreeSet<>();
            }
        } else {
            groupByValues = new LinkedHashSet<>();
        }
        for (Object bean : collection) {
            groupByValues.add(getGroupKey(bean, groupProperty));
        }
        for (Iterator<Object> iterator = groupByValues.iterator(); iterator.hasNext();) {
            Object groupValue = iterator.next();
            List<Object> groupItems = new ArrayList<>();
            for (Object bean : collection) {
                if (groupValue.equals(getGroupKey(bean, groupProperty))) {
                    groupItems.add(bean);
                }
            }
            if (!groupItems.isEmpty()) {
                result.add(new GroupData(groupItems.get(0), groupItems));
            }
        }
        return result;
    }

    /**
     * Groups items from an iterable collection using passed group property and group order
     * @param iterable iterable object
     * @param groupProperty property to use to group the items
     * @param groupOrder an order to sort the groups
     * @return a collection of group data objects
     */
    public static Collection<GroupData> groupIterable(Iterable<?> iterable, String groupProperty, String groupOrder) {
        Collection<GroupData> result = new ArrayList<GroupData>();
        if (iterable == null) {
            return result;
        }
        Set<Object> groupByValues;
        if (groupOrder != null) {
            if ("desc".equalsIgnoreCase(groupOrder)) {
                groupByValues = new TreeSet<>(Collections.reverseOrder());
            } else {
                groupByValues = new TreeSet<>();
            }
        } else {
            groupByValues = new LinkedHashSet<>();
        }
        for (Object bean : iterable) {
            groupByValues.add(getGroupKey(bean, groupProperty));
        }
        for (Iterator<Object> iterator = groupByValues.iterator(); iterator.hasNext();) {
            Object groupValue = iterator.next();
            List<Object> groupItems = new ArrayList<>();
            for (Object bean : iterable) {
                if (groupValue.equals(getGroupKey(bean, groupProperty))) {
                    groupItems.add(bean);
                }
            }
            if (!groupItems.isEmpty()) {
                result.add(new GroupData(groupItems.get(0), groupItems));
            }
        }
        return result;
    }

    private static Object getGroupKey(Object bean, String propertyName) {
        Object ret = getObjectProperty(bean, propertyName, true);
        return ret == null ? "null" : ret;
    }

    /**
     * Reads all the data from the input stream, and returns the bytes read.
     * 
     * @param stream -
     * @return byte array
     * @throws IOException -
     */
    public static byte[] toByteArray(InputStream stream) throws IOException {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        byte[] buffer = new byte[4096];
        int read = 0;
        while (read != -1) {
            read = stream.read(buffer);
            if (read > 0) {
                baos.write(buffer, 0, read);
            }
        }
        return baos.toByteArray();
    }

    /**
     * Evaluates passed collection name into a {@link Collection} object
     * @param expressionEvaluator
     * @param collectionName
     * @param context
     * @return an evaluated {@link Collection} instance
     */
    public static Collection<?> transformToCollectionObject(ExpressionEvaluator expressionEvaluator, String collectionName, Context context) {
        Object collectionObject = expressionEvaluator.evaluate(collectionName, context.toMap());
        if (!(collectionObject instanceof Collection)) {
            throw new JxlsException(collectionName + " expression is not a collection");
        }
        return (Collection<?>) collectionObject;
    }

    /**
     * @param cellRefEntry
     * @return the sheet name regular expression string
     */
    public static String sheetNameRegex(Map.Entry<CellRef, List<CellRef>> cellRefEntry) {
        return (cellRefEntry.getKey().isIgnoreSheetNameInFormat() ? "(?<!!)" : "");
    }

    /**
     * Creates a list of target formula cell references
     * @param targetFormulaCellRef
     * @param targetCells
     * @param cellRefsToExclude
     * @return
     */
    public static List<CellRef> createTargetCellRefListByColumn(CellRef targetFormulaCellRef, List<CellRef> targetCells,
            List<CellRef> cellRefsToExclude) {
        List<CellRef> resultCellList = new ArrayList<>();
        int col = targetFormulaCellRef.getCol();
        for (CellRef targetCell : targetCells) {
            if (targetCell.getCol() == col
                    && targetCell.getRow() < targetFormulaCellRef.getRow()
                    && !cellRefsToExclude.contains(targetCell)) {
                resultCellList.add(targetCell);
            }
        }
        return resultCellList;
    }

    /**
     * Calculates a number of occurences of a symbol in the string
     * @param str
     * @param ch
     * @return
     */
    public static int countOccurences(String str, char ch) {
        int count = 0;
        for (int i = 0; i < str.length(); i++) {
            if (str.charAt(i) == ch) {
                count++;
            }
        }
        return count;
    }

    /**
     * Evaluates the passed collection name into an {@link Iterable} object
     * @param expressionEvaluator
     * @param collectionName
     * @param context
     * @return an iterable object from the {@link Context} under given name
     */
    public static Iterable<Object> transformToIterableObject(ExpressionEvaluator expressionEvaluator, String collectionName, Context context) {
        Object collectionObject = expressionEvaluator.evaluate(collectionName, context.toMap());
        if (!(collectionObject instanceof Iterable)) {
            throw new JxlsException(collectionName + " expression is not a collection");
        }
        List<Object> ret = new ArrayList<>();
        for (Object i : (Iterable) collectionObject) {
            ret.add(i);
        }
        return ret;
    }

    /**
     * @param name
     * @return regular expression to detect the passed cell name
     */
    public static String getStrictCellNameRegex(String name) {
        return "(?<=[^A-Z]|^)" + name + "(?=\\D|$)";
    }

    /**
     * Return names of all multi sheet template
     *
     * @param areaList list of area
     * @return string array
     */
    static List<String> getSheetsNameOfMultiSheetTemplate(List<Area> areaList) {
        List<String> templateSheetsName = new ArrayList<>();
        for (Area xlsArea : areaList) {
            for (Command command : xlsArea.findCommandByName("each")) {
                boolean isAreaHasMultiSheetAttribute = ((EachCommand) command).getMultisheet() != null && !((EachCommand) command).getMultisheet().isEmpty();
                if (isAreaHasMultiSheetAttribute) {
                    templateSheetsName.add(xlsArea.getAreaRef().getSheetName());
                    break;
                }
            }
        }
        return templateSheetsName;
    }
}
