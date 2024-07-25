package org.jxls.functions;

import java.lang.reflect.InvocationTargetException;
import java.util.Map;

import org.apache.commons.beanutils2.PropertyUtils;
import org.jxls.command.RunVar;
import org.jxls.common.JxlsException;
import org.jxls.common.NeedsPublicContext;
import org.jxls.common.PublicContext;

/**
 * Group sum
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
public class GroupSum<T> implements NeedsPublicContext {
    private PublicContext context;
    private final SummarizerBuilder<T> sumBuilder;
    private String objectVarName = "i";
    
    public GroupSum(SummarizerBuilder<T> sumBuilder) {
        this.sumBuilder = sumBuilder;
    }

    @Override
    public void setPublicContext(PublicContext context) {
        this.context = context;
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
        Summarizer<T> sum = sumBuilder.build();
        try (RunVar runVar = new RunVar(objectVarName, context)) {
            for (Object object : collection) {
                runVar.put(object);
                if (context.isConditionTrue(filter)) {
                    sum.add(getValue(object, fieldName));
                }
            }
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
        return context.evaluate(expression);
    }
}
