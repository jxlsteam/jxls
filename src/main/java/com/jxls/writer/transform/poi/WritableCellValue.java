package com.jxls.writer.transform.poi;

import com.jxls.writer.common.Context;
import org.apache.poi.ss.usermodel.Cell;

/**
 * Defines an interface for a cell value which knows how to write itself to a cell
 * @author Leonid Vysochyn
 *         Date: 6/18/12
 */
public interface WritableCellValue {
    Object writeToCell(Cell cell, Context context);
}
