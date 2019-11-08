package org.jxls.util;

import org.jxls.common.JxlsException;

/**
 * JXLS exception: Excel file cannot be opened
 * 
 * <p>If this exception occurs:</p><ol>
 * <li>Check the filename.</li>
 * <li>Check the file format. (.xlsx files are Zip files.)</li>
 * <li>Check whether you can open the Excel file with Microsoft Excel.</li>
 * <li>Probably you must recreate the Excel file from scratch. Use of Microsoft Excel and XLSX format is recommended.</li></ol>
 */
public class CannotOpenWorkbookException extends JxlsException {
    private static final long serialVersionUID = -3618771481378341600L;

    public CannotOpenWorkbookException(Throwable e) {
        super(e);
    }
}
