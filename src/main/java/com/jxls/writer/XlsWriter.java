package com.jxls.writer;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.InputStream;
import java.io.OutputStream;

/**
 * @author Leonid Vysochyn
 *         Date: 25.04.2009
 */
public interface XlsWriter {
    public void transformTemplate(InputStream inputStream, OutputStream outputStream, BeanContext beanContext);
}
