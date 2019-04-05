package org.jxls.command;

import java.util.List;

import org.jxls.common.CellRef;
import org.jxls.common.Context;

/**
 * Creates cell references based on passed sheet names
 */
public class SheetNameGenerator implements CellRefGenerator {
    private final List<String> sheetNames;
    private final CellRef startCellRef;

    public SheetNameGenerator(List<String> sheetNames, CellRef startCellRef) {
        this.sheetNames = sheetNames;
        this.startCellRef = startCellRef;
    }

    @Override
    public CellRef generateCellRef(int index, Context context) {
        if (sheetNames.size() <= index) {
            return null;
        }
        String sheetName = sheetNames.get(index);
        return new CellRef(sheetName, startCellRef.getRow(), startCellRef.getCol());
    }
}
