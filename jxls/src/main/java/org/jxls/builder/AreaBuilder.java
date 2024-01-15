package org.jxls.builder;

import java.util.List;

import org.jxls.area.Area;
import org.jxls.transform.Transformer;

/**
 * Interface for building the areas
 */
public interface AreaBuilder {

    /**
     * @param transformer -
     * @param clearTemplateCells false: template cells will not be cleared if the expression can not be evaluated
     * @return areas
     */
    List<Area> build(Transformer transformer, boolean clearTemplateCells);
}
