package org.jxls.util;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;

import org.jxls.common.JxlsException;

/**
 * <p>An <code>OrderByComparator</code> is a <code>Comparator</code> that is
 * capable of comparing two objects based on a dynamic list of properties of
 * the objects of type <code>T</code>.  It can sort any of its properties
 * ascending or descending, and for any of its properties, it can place nulls
 * first or last.  Like SQL, this will default to ascending.  Nulls default to
 * last if ascending, and first if descending.</p>
 */
public class OrderByComparator<T> implements Comparator<T> {
    private UtilWrapper util;

    /**
     * Sort ascending (default).
     */
    public static final String ASC = "ASC";
    /**
     * Sort descending.
     */
    public static final String DESC = "DESC";

    /**
     * Constant to order ascending.
     */
    public static final int ORDER_ASC = 1;
    /**
     * Constant to order descending.
     */
    public static final int ORDER_DESC = -1;

    private List<String> myProperties;
    private List<Integer> myOrderings;
    private int mySize;

    /**
     * Constructs an <code>OrderByComparator</code> based on a <code>List</code>
     * of expressions, of the format "property [ASC|DESC]".
     * @param expressions A <code>List</code> of expressions.
     * @throws ParseException If there is a problem parsing the expressions.
     */
    public OrderByComparator(List<String> expressions, UtilWrapper util) {
        this.util = util;
        setExpressions(expressions);
    }

    /**
     * Sets the internal lists for all properties, order sequences, and null
     * order sequences.
     * @param expressions A <code>List</code> of expressions.
     * @throws ParseException If there is a problem parsing the expressions.
     */
    private void setExpressions(List<String> expressions) {
        if (expressions == null || expressions.size() <= 0) {
            throw new JxlsException("No order by expressions found.");
        }
        mySize = expressions.size();
        myProperties = new ArrayList<>(mySize);
        myOrderings = new ArrayList<>(mySize);
        for (String expr : expressions) {
            String[] parts = expr.split("\\s+");
            String property;
            int ordering;
            if (parts.length > 0 && parts.length < 5) {
                property = parts[0];
                ordering = ORDER_ASC;

                if (parts.length == 2 || parts.length == 4) {
                    // ordering is next.
                    if (ASC.equalsIgnoreCase(parts[1])) {
                        ordering = ORDER_ASC;
                    }
                    else if (DESC.equalsIgnoreCase(parts[1])) {
                        ordering = ORDER_DESC;
                    }
                    else {
                        throw new JxlsException("Expected \"" + ASC + "\" or \"" + DESC + ": " + expr);
                    }
                }
            }
            else {
                throw new JxlsException("Expected \"property\" [" + ASC + "|" + DESC + "] : " + expr);
            }

            myProperties.add(property);
            myOrderings.add(ordering);
        }
    }

    /**
     * <p>Compares the given objects to determine order.  Fulfills the
     * <code>Comparator</code> contract by returning a negative integer, 0, or a
     * positive integer if <code>o1</code> is less than, equal to, or greater
     * than <code>o2</code>.</p>
     * <p>This compare method respects all properties, their order sequences,
     * and their null order sequences.</p>
     *
     * @param o1 The left-hand-side object to compare.
     * @param o2 The right-hand-side object to compare.
     * @return A negative integer, 0, or a positive integer if <code>o1</code>
     *    is less than, equal to, or greater than <code>o2</code>.
     * @throws UnsupportedOperationException If any property specified in the
     *    constructor doesn't correspond to a no-argument "get&lt;Property&gt;"
     *    getter method in <code>T</code>, or if the property's type is not
     *    <code>Comparable</code>.
     */
    @SuppressWarnings({ "rawtypes", "unchecked" })
    @Override
    public int compare(T o1, T o2) throws UnsupportedOperationException {
        int comp = 0;
        for (int i = 0; i < mySize; i++) {
            String property = myProperties.get(i);
            int ordering = myOrderings.get(i);

            Comparable value1, value2;
            try {
                value1 = (Comparable) util.getObjectProperty(o1, property);
                value2 = (Comparable) util.getObjectProperty(o2, property);
            }
            catch (Exception e) {
                throw new JxlsException("No matching method found for \"" + property + "\".", e);
            }
            try {
                if (value1 == null) {
                    if (value2 == null)
                        comp = 0;
                }
                else {
                    comp = ordering * value1.compareTo(value2);
                }
                if (comp != 0) {
                    return comp;
                }
            }
            catch (ClassCastException e) {
                throw new UnsupportedOperationException("Property \"" + property + "\" needs to be Comparable.");
            }
        }
        return 0;
    }

    /**
     * Indicates whether the given <code>OrderByComparator</code> is equal to
     * this <code>OrderByComparator</code>.  All property names must match in
     * order, and all of the order sequences and null order sequences must
     * match.
     *
     * @param obj The other <code>OrderByComparator</code>.
     */
    @Override
    public boolean equals(Object obj) {
        if (obj instanceof OrderByComparator) {
            OrderByComparator otherComp = (OrderByComparator) obj;
            if (mySize != otherComp.mySize) {
                return false;
            }
            for (int i = 0; i < mySize; i++) {
                if (!myProperties.get(i).equals(otherComp.myProperties.get(i)))
                    return false;
                if (myOrderings.get(i) != otherComp.myOrderings.get(i)) {
                    return false;
                }
            }
            return true;
        }
        return false;
    }

    /**
     * Returns a <code>List</code> of all properties.
     * @return A <code>List</code> of all properties.
     */
    public List<String> getProperties() {
        return myProperties;
    }

    /**
     * Returns a <code>List</code> of orderings.
     * @return A <code>List</code> of orderings.
     * @see #ORDER_ASC
     * @see #ORDER_DESC
     */
    public List<Integer> getOrderings() {
        return myOrderings;
    }
}
