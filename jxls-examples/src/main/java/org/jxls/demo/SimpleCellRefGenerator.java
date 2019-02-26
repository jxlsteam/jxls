package org.jxls.demo;

import org.jxls.command.CellRefGenerator;
import org.jxls.common.CellRef;
import org.jxls.common.Context;

/**
 * @author Leonid Vysochyn
 */
public class SimpleCellRefGenerator implements CellRefGenerator {
    public CellRef generateCellRef(int index, Context context) {
        return new CellRef("sheet" + index + "!B2");
    }
}
