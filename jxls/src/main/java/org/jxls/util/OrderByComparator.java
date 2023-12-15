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
    private final UtilWrapper util;

    /**
     * Sort ascending (default).
     */
    public static final String ASC = "ASC";
    public static final String ASC_IGNORECASE = "ASC_IGNORECASE";
    /**
     * Sort descending.
     */
    public static final String DESC = "DESC";
    public static final String DESC_IGNORECASE = "DESC_IGNORECASE";

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
    private List<Boolean> myIgnoreCase;
    private int mySize;

    /**
     * Constructs an <code>OrderByComparator</code> based on a <code>List</code>
     * of expressions, of the format "property [ASC|DESC]".
     * @param expressions A <code>List</code> of expressions.
     * @param util -
     * @throws JxlsException If there is a problem parsing the expressions.
     */
    public OrderByComparator(List<String> expressions, UtilWrapper util) {
        this.util = util;
        setExpressions(expressions);
    }

    /**
     * Sets the internal lists for all properties, order sequences, and null
     * order sequences.
     * @param expressions A <code>List</code> of expressions.
     * @throws JxlsException If there is a problem parsing the expressions.
     */
    private void setExpressions(List<String> expressions) {
        if (expressions == null || expressions.size() <= 0) {
            throw new JxlsException("No order by expressions found.");
        }
        mySize = expressions.size();
        myProperties = new ArrayList<>(mySize);
        myOrderings = new ArrayList<>(mySize);
        myIgnoreCase = new ArrayList<>(mySize);
        for (String expr : expressions) {
            String[] parts = expr.trim().split("\\s+");
            String property;
            int ordering;
            Boolean ignoreCase;
            if (parts.length > 0 && parts.length < 5) {
                property = parts[0];
                ordering = ORDER_ASC;
                ignoreCase = Boolean.FALSE;

                if (parts.length == 2 || parts.length == 4) {
                    // ordering is next.
                    if (ASC.equalsIgnoreCase(parts[1])) {
                        ordering = ORDER_ASC;
                    } else if (ASC_IGNORECASE.equalsIgnoreCase(parts[1])) {
                        ordering = ORDER_ASC;
                        ignoreCase = Boolean.TRUE;
                    } else if (DESC.equalsIgnoreCase(parts[1])) {
                        ordering = ORDER_DESC;
                    } else if (DESC_IGNORECASE.equalsIgnoreCase(parts[1])) {
                        ordering = ORDER_DESC;
                        ignoreCase = Boolean.TRUE;
                    } else {
						throw new JxlsException("Expected \"" + ASC + "\", \"" + DESC + "\", \"" + ASC_IGNORECASE
								+ "\" or \"" + DESC_IGNORECASE + "\": " + expr);
                    }
                }
            } else {
				throw new JxlsException("Expected \"property\" [" + ASC + "|" + DESC + "|" + ASC_IGNORECASE + "|"
						+ DESC_IGNORECASE + "] : " + expr);
            }

            myProperties.add(property);
            myOrderings.add(ordering);
            myIgnoreCase.add(ignoreCase);
        }
    }

    @SuppressWarnings({ "rawtypes", "unchecked" })
    @Override
    public int compare(T o1, T o2) throws UnsupportedOperationException {
        int comp = 0;
        for (int i = 0; i < mySize; i++) {
            String property = myProperties.get(i);
            int ordering = myOrderings.get(i);
            boolean ignoreCase = myIgnoreCase.get(i).booleanValue();

            Comparable value1, value2;
            try {
                value1 = (Comparable) util.getObjectProperty(o1, property);
                value2 = (Comparable) util.getObjectProperty(o2, property);
            } catch (ClassCastException e) {
                throw new JxlsException("Property \"" + property + "\" must implement Comparable.");
            } catch (Exception e) {
                throw new JxlsException("Error accessing property \"" + property + "\".", e);
            }
            if (value1 != null && value2 != null) {
            	if (ignoreCase) {
					if (value1 instanceof String && value2 instanceof String) {
						comp = ((String) value1).compareToIgnoreCase((String) value2) * ordering;
					} else {
						throw new JxlsException("Property \"" + property + "\" must be a String if you use "
								+ ASC_IGNORECASE + " or " + DESC_IGNORECASE + ".");
					}
            	} else {
            		comp = value1.compareTo(value2) * ordering;
            	}
                if (comp != 0) {
                    return comp;
                } // else: continue with next sort attribute
            } else if (value1 == null && value2 != null) {
                return 1 * ordering;
            } else if (value1 != null && value2 == null) {
                return -1 * ordering;
            }
        }
        return 0;
    }
}
