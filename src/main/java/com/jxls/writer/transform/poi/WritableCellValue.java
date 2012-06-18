package com.jxls.writer.transform.poi;

import com.jxls.writer.common.Context;
import org.apache.poi.ss.usermodel.Cell;

/**
 * @author Leonid Vysochyn
 *         Date: 6/18/12 4:34 PM
 */
public interface WritableCellValue {
    Object writeToCell(Cell cell, Context context);
}
