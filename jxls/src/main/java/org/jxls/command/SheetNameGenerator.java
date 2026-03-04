package org.jxls.command;

import java.util.HashSet;
import java.util.List;
import java.util.Set;

import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.logging.JxlsLogger;
import org.jxls.transform.SafeSheetNameBuilder;

/**
 * Creates cell references based on passed sheet names
 */
public class SheetNameGenerator implements CellRefGenerator {
    public static final String ERR_MSG = "Duplicate sheet name ";
    private final List<String> sheetNames;
    private final CellRef startCellRef;
    private final Set<String> usedSheetNames = new HashSet<>(); // only used if there's no SafeSheetNameBuilder

    public SheetNameGenerator(List<String> sheetNames, CellRef startCellRef) {
        this.sheetNames = sheetNames;
        this.startCellRef = startCellRef;
    }

    @Override
    public CellRef generateCellRef(int index, Context context, JxlsLogger logger) {
        String sheetName = index >= 0 && index < sheetNames.size() ? sheetNames.get(index) : null;
        Object builder = RunVar.getRunVar(SafeSheetNameBuilder.CONTEXT_VAR_NAME, context);
        if (builder instanceof SafeSheetNameBuilder b) {
            sheetName = b.createSafeSheetName(sheetName, index, logger);
        } else if (!usedSheetNames.add(sheetName)) {
            logger.error(ERR_MSG + sheetName);
            return null;
        }
        if (sheetName == null) {
            return null;
        }
        return new CellRef(sheetName, startCellRef.getRow(), startCellRef.getCol());
    }
}
