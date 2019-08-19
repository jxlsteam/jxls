package org.jxls.command;

import java.util.List;

import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.apache.poi.ss.util.WorkbookUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Creates cell references based on passed sheet names
 */
public class SheetNameGenerator implements CellRefGenerator {
    private final List<String> sheetNames;
    private final CellRef startCellRef;
    private static Logger logger = LoggerFactory.getLogger(SheetNameGenerator.class);

    public SheetNameGenerator(List<String> sheetNames, CellRef startCellRef) {
        this.sheetNames = sheetNames;
        this.startCellRef = startCellRef;
    }

    @Override
    public CellRef generateCellRef(int index, Context context) {
        if (sheetNames.size() <= index) {
            return null;
        }

        String givenSheetName = sheetNames.get(index);
        String sheetName = WorkbookUtil.createSafeSheetName(givenSheetName);
        if (givenSheetName != sheetName) {
            logger.warn("Given sheet name was invalid, updating. originalSheetName={} newSheetName={}", givenSheetName, sheetName);
        }
        return new CellRef(sheetName, startCellRef.getRow(), startCellRef.getCol());
    }
}
