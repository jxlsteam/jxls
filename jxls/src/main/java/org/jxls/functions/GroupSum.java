package org.jxls.functions;

import java.lang.reflect.InvocationTargetException;
import java.util.Collection;
import java.util.Map;

import org.apache.commons.beanutils.PropertyUtils;
import org.jxls.common.Context;
import org.jxls.expression.JexlExpressionEvaluator;

/**
 * <h1>Group sum</h1>
 * <p>The sum function for calculation a group sum takes two arguments: the collection as JEXL expression (or its name as a String)
 * and the name (as String) of the attribute. The attribute can be a object property or a Map entry. The value type T can be of any
 * type and is implemented by a generic SummarizerBuilder.</p>
 * 
 * <h2>Example</h2>
 * <p>Add an instance of this class e.g. with name "G" to your Context.</p>
 * <pre>${G.sum("salary", employees.items)}</pre>
 * <p>Above the 2nd argument is a JEXL expression. The collection name as String is also possible:</p>
 * <pre>${G.sum("salary", "employees.items")}</pre>
 */
public class GroupSum<T> {
    private final Context context;
    private final SummarizerBuilder<T> sumBuilder;
    
    public GroupSum(Context context, SummarizerBuilder<T> sumBuilder) {
        this.context = context;
        this.sumBuilder = sumBuilder;
    }
    
    /**
     * Returns the sum of the given field of all items.
     * 
     * @param fieldName name of the field of type T to be summed (without the loop var name!)
     * @param expression JEXL expression as String, usually name of the Collection, often ends with ".items"
     * @return sum of type T
     */
    public T sum(String fieldName, String expression) {
        return sum(fieldName, getItems(expression));
    }
    
    /**
     * Returns the sum of the given field of all items.
     * 
     * @param fieldName name of the field of type T to be summed (without the loop var name!)
     * @param collection the collection; inside Excel file it's a JEXL expression
     * @return sum of type T
     */
    public T sum(String fieldName, Collection<Object> collection) {
        Summarizer<T> sum = sumBuilder.build();
        for (Object i : collection) {
            sum.add(getValue(i, fieldName));
        }
        return sum.getSum();
    }
    
    private Object getValue(Object i, String fieldName) {
        if (i instanceof Map) {
            Map<?,?> map = (Map<?,?>) i;
            if (!map.containsKey(fieldName)) {
                throw new RuntimeException("Attribute " + fieldName + " does not exist in collection element!");
            }
            return map.get(fieldName);
        } else {
            try {
                return PropertyUtils.getProperty(i, fieldName);
            } catch (IllegalAccessException | InvocationTargetException | NoSuchMethodException e) {
                throw new RuntimeException(e);
            }
        }
    }

    @SuppressWarnings("unchecked")
    private Collection<Object> getItems(String expression) {
        Object result = getValue(expression);
        if (result == null) {
            throw new NullPointerException("\"" + expression + "\" is null!");
        } else if (!(result instanceof Collection)) {
            throw new ClassCastException(expression + " is not a Collection!");
        }
        return (Collection<Object>) result;
    }

    private Object getValue(String expression) {
        return new JexlExpressionEvaluator(expression).evaluate(context.toMap());
    }
}
