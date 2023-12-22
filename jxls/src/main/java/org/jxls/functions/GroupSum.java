package org.jxls.functions;

import java.lang.reflect.InvocationTargetException;
import java.util.Map;

import org.apache.commons.beanutils.PropertyUtils;
import org.jxls.common.Context;
import org.jxls.common.JxlsException;
import org.jxls.expression.ExpressionEvaluator;
import org.jxls.expression.JexlExpressionEvaluator;
import org.jxls.transform.TransformationConfig;

/**
 * Group sum
 * <p>The sum function for calculation a group sum takes two arguments: the collection as JEXL expression (or its name as a String)
 * and the name (as String) of the attribute. The attribute can be a object property or a Map entry. The value type T can be of any
 * type and is implemented by a generic SummarizerBuilder.</p>
 * <p>Call setTransformationConfig() if you want to use the methods with filter condition parameter.</p>
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
    private TransformationConfig transformationConfig;
    private String objectVarName = "i";
    
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
    public T sum(String fieldName, Iterable<Object> collection) {
        Summarizer<T> sum = sumBuilder.build();
        for (Object i : collection) {
            sum.add(getValue(i, fieldName));
        }
        return sum.getSum();
    }

    public TransformationConfig getTransformationConfig() {
        return transformationConfig;
    }

    public void setTransformationConfig(TransformationConfig transformationConfig) {
        this.transformationConfig = transformationConfig;
    }

    public String getObjectVarName() {
        return objectVarName;
    }

    public void setObjectVarName(String objectVarName) {
        this.objectVarName = objectVarName;
    }

    public T sum(String fieldName, String expression, String filter) {
        return sum(fieldName, getItems(expression), filter);
    }
    
    public T sum(String fieldName, Iterable<Object> collection, String filter) {
        if (transformationConfig == null) {
            throw new JxlsException("Please set GroupSum.transformationConfig!");
        }
        ExpressionEvaluator expressionEvaluator = transformationConfig.getExpressionEvaluator();
        Summarizer<T> sum = sumBuilder.build();
        Object oldValue = context.getRunVar(objectVarName);
        for (Object i : collection) {
            Object value = getValue(i, fieldName);
            context.putVar(objectVarName, i);
            if (Boolean.TRUE.equals(expressionEvaluator.isConditionTrue(filter, context.toMap()))) {
                sum.add(value);
            }
        }
        if (oldValue != null) {
            context.putVar(objectVarName, oldValue); // TODO runVar support
        } else {
            context.removeVar(objectVarName);
        }
        return sum.getSum();
    }

    private Object getValue(Object i, String fieldName) {
        if (i instanceof Map<?,?> map) {
            if (!map.containsKey(fieldName)) {
                throw new JxlsException("Attribute " + fieldName + " does not exist in collection element!");
            }
            return map.get(fieldName);
        } else {
            try {
                return PropertyUtils.getProperty(i, fieldName);
            } catch (IllegalAccessException | InvocationTargetException | NoSuchMethodException e) {
                throw new JxlsException(e);
            }
        }
    }

    private Iterable<Object> getItems(String expression) {
        Object result = getValue(expression);
        if (result == null) {
            throw new NullPointerException("\"" + expression + "\" is null!");
        } else if (!(result instanceof Iterable)) {
            throw new ClassCastException(expression + " is not an Iterable!");
        }
        return (Iterable<Object>) result;
    }

    private Object getValue(String expression) {
        if (transformationConfig == null) {
            return new JexlExpressionEvaluator(expression).evaluate(context.toMap()); // TODO not good, but how to get TransformerConfig? Yes I know there's setTransformationConfig()
        }
        return transformationConfig.getExpressionEvaluator().evaluate(expression, context.toMap());
    }
}
