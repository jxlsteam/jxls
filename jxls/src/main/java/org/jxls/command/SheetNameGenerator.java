package org.jxls.command;

import java.util.List;

import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.transform.SafeSheetNameBuilder;

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
        String sheetName = index >= 0 && index < sheetNames.size() ? sheetNames.get(index) : null;
        Object builder = context.getVar(SafeSheetNameBuilder.CONTEXT_VAR_NAME);
        if (builder instanceof SafeSheetNameBuilder) {
            sheetName = ((SafeSheetNameBuilder) builder).createSafeSheetName(sheetName, index);
        }
        if (sheetName == null) {
            return null;
        }
        return new CellRef(sheetName, startCellRef.getRow(), startCellRef.getCol());
    }
}
