package com.jxls.writer.poi;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.InputStream;
import java.io.OutputStream;

import com.jxls.writer.XlsWriter;
import com.jxls.writer.BeanContext;

/**
 * @author Leonid Vysochyn
 *         Date: 25.04.2009
 */
public class PoiXlsWriter implements XlsWriter {
    public void transformTemplate(InputStream inputStream, OutputStream outputStream, BeanContext beanContext) {
        try {
            Workbook workbook = WorkbookFactory.create( inputStream );
            
            workbook.write( outputStream );
        } catch (Exception e) {
            throw new RuntimeException("Error while transforming template", e);
        }
    }
}
