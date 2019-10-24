package org.jxls.util;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.util.Collection;
import java.util.List;
import java.util.Map;

import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.GroupData;
import org.jxls.expression.ExpressionEvaluator;

/**
 * A wrapper around {@link Util} class to prevent static methods spreading
 */
public class UtilWrapper {

    public List<String> getFormulaCellRefs(String formula) {
        return Util.getFormulaCellRefs(formula);
    }

    public List<String> getJointedCellRefs(String formula) {
        return Util.getJointedCellRefs(formula);
    }

    public List<String> getCellRefsFromJointedCellRef(String jointedCellRef) {
        return Util.getCellRefsFromJointedCellRef(jointedCellRef);
    }

    public boolean formulaContainsJointedCellRef(String formula) {
        return Util.formulaContainsJointedCellRef(formula);
    }

    public String createTargetCellRef(List<CellRef> targetCellDataList) {
        return Util.createTargetCellRef(targetCellDataList);
    }

    public String joinStrings(List<String> strings, String separator) {
        return Util.joinStrings(strings, separator);
    }

    public List<List<CellRef>> groupByRanges(List<CellRef> cellRefList, int targetRangeCount) {
        return Util.groupByRanges(cellRefList, targetRangeCount);
    }

    public List<List<CellRef>> groupByColRange(List<CellRef> cellRefList) {
        return Util.groupByColRange(cellRefList);
    }

    public List<List<CellRef>> groupByRowRange(List<CellRef> cellRefList) {
        return groupByRowRange(cellRefList);
    }

    public Boolean isConditionTrue(ExpressionEvaluator evaluator, String condition, Context context) {
        return Util.isConditionTrue(evaluator, condition, context);
    }

    public Boolean isConditionTrue(ExpressionEvaluator evaluator, Context context) {
        return Util.isConditionTrue(evaluator, context);
    }

    public void setObjectProperty(Object obj, String propertyName, String propertyValue, boolean ignoreNonExisting) {
        Util.setObjectProperty(obj, propertyName, propertyValue, ignoreNonExisting);
    }

    public void setObjectProperty(Object obj, String propertyName, String propertyValue) throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {
        Util.setObjectProperty(obj, propertyName, propertyValue);
    }

    public Object getObjectProperty(Object obj, String propertyName, boolean failSilently) {
        return Util.getObjectProperty(obj, propertyName, failSilently);
    }

    public Object getObjectProperty(Object obj, String propertyName) throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {
        return Util.getObjectProperty(obj, propertyName);
    }

    public Collection<GroupData> groupCollection(Collection<?> collection, String groupProperty, String groupOrder) {
        return Util.groupCollection(collection, groupProperty, groupOrder);
    }

    public Collection<GroupData> groupIterable(Iterable<?> iterable, String groupProperty, String groupOrder) {
        return Util.groupIterable(iterable, groupProperty, groupOrder);
    }

    public byte[] toByteArray(InputStream stream) throws IOException {
        return Util.toByteArray(stream);
    }

    public Collection<?> transformToCollectionObject(ExpressionEvaluator expressionEvaluator, String collectionName, Context context) {
        return Util.transformToCollectionObject(expressionEvaluator, collectionName, context);
    }

    public Iterable<Object> transformToIterableObject(ExpressionEvaluator expressionEvaluator, String collectionName,
            Context context) {
        return (Iterable<Object>) Util.transformToIterableObject(expressionEvaluator, collectionName, context);
    }

    public String sheetNameRegex(Map.Entry<CellRef, List<CellRef>> cellRefEntry) {
        return Util.sheetNameRegex(cellRefEntry);
    }

    public List<CellRef> createTargetCellRefListByColumn(CellRef targetFormulaCellRef, List<CellRef> targetCells, List<CellRef> cellRefsToExclude) {
        return Util.createTargetCellRefListByColumn(targetFormulaCellRef, targetCells, cellRefsToExclude);
    }
}
