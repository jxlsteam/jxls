package org.jxls.formula;

import org.jxls.area.Area;
import org.jxls.transform.Transformer;

/**
 * Defines formula processing for {@link Area}
 */
public interface FormulaProcessor {

    /**
     * Processes all formulas using the provided transformer
     * Used after area transformation
     * Deprecated - use processAreaFormulas(Transformer transformer, Area area) method instead
     * @param transformer -
     */
    @Deprecated
    void processAreaFormulas(Transformer transformer);

    /**
     * Processes all
     * @param transformer transformer to use for formula processing
     * @param area - xls area for which the formula processing is invoked
     */
    void processAreaFormulas(Transformer transformer, Area area);
}
