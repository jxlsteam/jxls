package com.jxls.writer.builder;

import com.jxls.writer.area.Area;

import java.io.InputStream;
import java.util.List;

/**
 * @author Leonid Vysochyn
 *         Date: 2/14/12 11:51 AM
 */
public interface AreaBuilder {
    List<Area> build(InputStream is);
}
