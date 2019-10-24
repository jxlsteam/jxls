package org.jxls.builder;

import org.jxls.area.Area;
import org.jxls.transform.Transformer;

import java.util.List;

/**
 * Interface to build an {@link Area}
 */
public interface AreaBuilder {

    List<Area> build();

    void setTransformer(Transformer transformer);

    Transformer getTransformer();
}
