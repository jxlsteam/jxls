package org.jxls.transformer;

import org.apache.poi.ss.usermodel.Workbook;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.templatebasedtests.UseOwnTransformerTest;
import org.jxls.transform.poi.PoiTransformer;

/**
 * Own implementation of PoiTransformer, always with streaming.
 * 
 * @see UseOwnTransformerTest
 */
public class OwnTransformer extends PoiTransformer {
    private boolean transformCalled = false;
    private boolean clearCellCalled = false;
    
    public OwnTransformer(Workbook workbook) {
        super(workbook, true);
    }
    
    // It must be possible to modify transform().
    @Override
    public void transform(CellRef srcCellRef, CellRef targetCellRef, Context context, boolean updateRowHeightFlag) {
        if (isStreaming()) { // Access to isSXSSF must be possible.
            transformCalled = true;
        }
        // ...
        super.transform(srcCellRef, targetCellRef, context, updateRowHeightFlag);
    }
    
    // It must be possible to modify clearCell().
    @Override
    public void clearCell(CellRef cellRef) {
        clearCellCalled = true;
        // ...
        super.clearCell(cellRef);
    }

    public boolean isTransformCalled() {
        return transformCalled;
    }

    public boolean isClearCellCalled() {
        return clearCellCalled;
    }
}
