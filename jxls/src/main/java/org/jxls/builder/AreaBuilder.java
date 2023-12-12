package org.jxls.builder;

import java.util.List;

import org.jxls.area.Area;
import org.jxls.transform.Transformer;

/**
 * Interface for building the areas
 */
public interface AreaBuilder {

    List<Area> build(Transformer transformer, boolean clearTemplateCells);
}
