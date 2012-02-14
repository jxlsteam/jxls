package com.jxls.writer.builder;

import com.jxls.writer.command.Area;

import java.io.InputStream;

/**
 * @author Leonid Vysochyn
 *         Date: 2/14/12 11:51 AM
 */
public interface AreaBuilder {
    Area build(InputStream is);
}
