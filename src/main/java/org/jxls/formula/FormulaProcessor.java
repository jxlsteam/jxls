package org.jxls.formula;

import org.jxls.area.Area;
import org.jxls.transform.Transformer;

/**
 * Defines formula processing for {@link Area}
 * Created by Leonid Vysochyn on 02-Aug-15.
 */
public interface FormulaProcessor {
    void processAreaFormulas(Transformer transformer);
}
