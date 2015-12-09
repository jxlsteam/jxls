package org.jxls.formula;

import org.jxls.area.Area;
import org.jxls.transform.Transformer;

/**
 * Defines formula processing for {@link Area}
 */
public interface FormulaProcessor {
    void processAreaFormulas(Transformer transformer);
}
