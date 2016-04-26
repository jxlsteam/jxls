package org.jxls.command;

import org.jxls.common.CellRef;
import org.jxls.common.Context;

import java.util.List;

/**
 * Creates cell references based on passed sheet names
 */
public class SheetNameGenerator implements CellRefGenerator {
    private List<String> sheetNames;
    private CellRef startCellRef;

    public SheetNameGenerator(List<String> sheetNames, CellRef startCellRef) {
        this.sheetNames = sheetNames;
        this.startCellRef = startCellRef;
    }

    @Override
    public CellRef generateCellRef(int index, Context context) {
        String sheetName = sheetNames.get(index);
        return new CellRef(sheetName, startCellRef.getRow(), startCellRef.getCol());
    }
}
