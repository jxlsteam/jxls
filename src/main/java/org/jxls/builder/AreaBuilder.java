package org.jxls.builder;

import org.jxls.area.Area;
import org.jxls.transform.Transformer;

import java.util.List;

/**
 * An area builder interface
 * @author Leonid Vysochyn
 *         Date: 2/14/12
 */
public interface AreaBuilder {
    List<Area> build();
    void setTransformer(Transformer transformer);
    Transformer getTransformer();
}
